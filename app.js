document.addEventListener('DOMContentLoaded', () => {
    const targetInput = document.getElementById('target-images');
    const directoryInput = document.getElementById('search-directory');
    const searchBtn = document.getElementById('search-btn');
    const targetPreview = document.getElementById('target-preview');
    const directoryInfo = document.getElementById('directory-info');
    const resultsSection = document.getElementById('results-section');
    const resultsGrid = document.getElementById('results-grid');
    const progressContainer = document.getElementById('progress-container');
    const progressText = document.getElementById('progress-text');
    const clearMemoryBtn = document.getElementById('clear-memory-btn');
    const saveProgressBtn = document.getElementById('save-progress-btn');
    const loadProgressBtn = document.getElementById('load-progress-btn');
    const loadProgressInput = document.getElementById('load-progress-input');
    const sheetImageInput = document.getElementById('sheet-image');
    const scanSheetBtn = document.getElementById('scan-sheet-btn');
    const sheetStatus = document.getElementById('sheet-status');
    const sheetSensitivityInput = document.getElementById('sheet-sensitivity');
    const sheetMaxPhotosInput = document.getElementById('sheet-max-photos');
    const sheetEnhanceInput = document.getElementById('sheet-enhance');
    
    const exportBar = document.getElementById('export-bar');
    const organizeSpreadsBtn = document.getElementById('organize-spreads-btn');
    const exportFolderBtn = document.getElementById('export-folder-btn');
    const selectionCount = document.getElementById('selection-count');
    const successModal = document.getElementById('export-success-modal');
    const closeModalBtn = document.getElementById('close-modal-btn');
    const closeModalX = document.getElementById('close-modal-x');

    const workspace = document.getElementById('spreads-workspace');
    const closeWorkspaceBtn = document.getElementById('close-workspace-btn');
    const spreadsContainer = document.getElementById('spreads-container');
    const imagesDock = document.getElementById('images-dock');
    const addSpreadBtn = document.getElementById('add-spread-btn');
    const workspaceExportPdfBtn = document.getElementById('workspace-export-pdf-btn');

    // Setup modal closing listeners
    [closeModalBtn, closeModalX].forEach(btn => {
        if (btn) {
            btn.addEventListener('click', () => {
                successModal.classList.add('hidden');
            });
        }
    });

    // State Variables
    let targetDescriptorsList = []; // Array of { file, keypoints, descriptors }
    let searchFiles = [];
    let isCvReady = window.cvReady || false;
    let selectedFiles = []; // Array of { file, card, badge }
    let draggedElement = null; // HTML element currently being dragged
    let pendingSheetImage = null;
    const featureCache = new Map(); // key => { keypoints, descriptors }

    /* ========================================================================= */
    /* ========================= INDEXEDDB MEMORY ============================== */
    /* ========================================================================= */
    
    const DB_NAME = 'ImageSearchMemory';
    const STORE_NAME = 'savedImages';

    function openDB() {
        return new Promise((resolve, reject) => {
            const request = indexedDB.open(DB_NAME, 1);
            request.onupgradeneeded = (e) => {
                e.target.result.createObjectStore(STORE_NAME, { keyPath: 'id' });
            };
            request.onsuccess = () => resolve(request.result);
            request.onerror = () => reject(request.error);
        });
    }

    async function saveToDB(file) {
        try {
            const db = await openDB();
            const tx = db.transaction(STORE_NAME, 'readwrite');
            const store = tx.objectStore(STORE_NAME);
            const buffer = await file.arrayBuffer();
            const blob = new Blob([buffer], { type: file.type || 'image/jpeg' });
            
            store.put({ 
                id: file.name, 
                blob: blob, 
                name: file.name, 
                path: file.webkitRelativePath || file.name 
            });
        } catch(e) { console.error("DB Save Error:", e); }
    }

    async function removeFromDB(id) {
        try {
            const db = await openDB();
            const tx = db.transaction(STORE_NAME, 'readwrite');
            tx.objectStore(STORE_NAME).delete(id);
        } catch(e) { console.error("DB Remove Error:", e); }
    }

    async function getAllFromDB() {
        try {
            const db = await openDB();
            return new Promise((resolve, reject) => {
                const tx = db.transaction(STORE_NAME, 'readonly');
                const req = tx.objectStore(STORE_NAME).getAll();
                req.onsuccess = () => resolve(req.result);
                req.onerror = reject;
            });
        } catch(e) { return []; }
    }

    async function clearDB() {
        try {
            const db = await openDB();
            const tx = db.transaction(STORE_NAME, 'readwrite');
            tx.objectStore(STORE_NAME).clear();
        } catch(e) { console.error("DB Clear Error:", e); }
    }

    async function restoreMemory() {
        try {
            const savedItems = await getAllFromDB();
            if (savedItems && savedItems.length > 0) {
                resultsSection.classList.remove('hidden');
                clearMemoryBtn.classList.remove('hidden');
                
                const divider = document.createElement('div');
                divider.className = 'search-divider';
                divider.innerHTML = `<div>Previous Session Memory</div> <span>Loaded securely from local database</span>`;
                resultsGrid.appendChild(divider);
                
                savedItems.forEach(item => {
                    const file = new File([item.blob], item.name, { type: item.blob.type });
                    Object.defineProperty(file, 'webkitRelativePath', { value: item.path });
                    
                    const card = displayMatch(file, "Saved Selection");
                    
                    // Force visually selected
                    card.classList.add('selected');
                    const badge = card.querySelector('.selection-badge');
                    selectedFiles.push({ file, card, badge });
                });
                
                updateExportBar();
            }
        } catch (e) {
            console.error("Failed to restore memory", e);
        }
    }

    // Attempt restore on boot
    restoreMemory();

    // Listen for OpenCV loading
    document.addEventListener('opencvReady', () => {
        isCvReady = true;
        checkReadyState();
        if (targetInput.files.length > 0) {
            processTargets(Array.from(targetInput.files));
        }
    });

    // Allows the UI to update unblock during heavy synchronous-like loops
    function yieldToMain() {
        return new Promise(resolve => setTimeout(resolve, 0));
    }

    // Helper: load file as an HTMLImageElement
    function loadImage(file) {
        return new Promise((resolve, reject) => {
            const reader = new FileReader();
            reader.onload = (e) => {
                const img = new Image();
                img.onload = () => resolve(img);
                img.onerror = () => reject(new Error('Failed to load image'));
                img.src = e.target.result;
            };
            reader.readAsDataURL(file);
        });
    }

    // Convert Image Element to cv.Mat
    function getMatFromImage(imgElement, maxDim = 600) {
        const canvas = document.createElement('canvas');
        canvas.width = imgElement.naturalWidth || imgElement.width || 300;
        canvas.height = imgElement.naturalHeight || imgElement.height || 300;
        
        // Scale down huge images to prevent freezing the browser
        let w = canvas.width;
        let h = canvas.height;
        if (w > maxDim || h > maxDim) {
            const ratio = Math.min(maxDim/w, maxDim/h);
            w = Math.floor(w * ratio);
            h = Math.floor(h * ratio);
        }
        canvas.width = w;
        canvas.height = h;

        const ctx = canvas.getContext('2d');
        ctx.drawImage(imgElement, 0, 0, w, h);
        
        try {
            return cv.imread(canvas);
        } catch (e) {
            console.error("OpenCV imread failed", e);
            return null;
        }
    }

    // Extract ORB Keypoints and Descriptors from an ImageElement
    // Includes grayscale normalization to improve lighting robustness.
    function extractFeatures(imgElement, options = {}) {
        const maxDim = options.maxDim || 600;
        const orbFeatures = options.orbFeatures || 500;

        const mat = getMatFromImage(imgElement, maxDim);
        if (!mat) return null;

        const gray = new cv.Mat();
        const normalized = new cv.Mat();
        cv.cvtColor(mat, gray, cv.COLOR_RGBA2GRAY);
        cv.equalizeHist(gray, normalized);
        
        const orb = new cv.ORB(orbFeatures); 
        const keypoints = new cv.KeyPointVector();
        const descriptors = new cv.Mat();
        const mask = new cv.Mat();

        try {
            orb.detectAndCompute(normalized, mask, keypoints, descriptors);
        } catch (e) {
            console.error("ORB compute failed", e);
        }
        
        // Memory cleanup
        gray.delete();
        normalized.delete();
        mask.delete();
        orb.delete();
        mat.delete();

        return { keypoints, descriptors };
    }

    function getFeatureCacheKey(file, options) {
        return [
            file.name,
            file.size,
            file.lastModified || 0,
            options.maxDim || 600,
            options.orbFeatures || 500
        ].join('::');
    }

    async function getOrCreateFileFeatures(file, options = {}) {
        const cacheKey = getFeatureCacheKey(file, options);
        if (featureCache.has(cacheKey)) {
            return featureCache.get(cacheKey);
        }

        const imgEl = await loadImage(file);
        const feats = extractFeatures(imgEl, options);
        if (feats && feats.descriptors && feats.descriptors.rows > 0) {
            featureCache.set(cacheKey, feats);
            return feats;
        }

        if (feats) {
            if (feats.descriptors) feats.descriptors.delete();
            if (feats.keypoints) feats.keypoints.delete();
        }
        return null;
    }

    async function getOrCreateHiResTargetFeatures(target) {
        if (target.hiResFeatures) return target.hiResFeatures;
        const feats = await getOrCreateFileFeatures(target.file, { maxDim: 1024, orbFeatures: 1200 });
        target.hiResFeatures = feats;
        return feats;
    }

    function clearFeatureCache() {
        featureCache.forEach(feats => {
            if (!feats) return;
            if (feats.descriptors && !feats.descriptors.isDeleted()) feats.descriptors.delete();
            if (feats.keypoints && !feats.keypoints.isDeleted()) feats.keypoints.delete();
        });
        featureCache.clear();
    }

    function orderCornerPoints(points) {
        const sums = points.map(p => p.x + p.y);
        const diffs = points.map(p => p.x - p.y);
        const topLeft = points[sums.indexOf(Math.min(...sums))];
        const bottomRight = points[sums.indexOf(Math.max(...sums))];
        const topRight = points[diffs.indexOf(Math.max(...diffs))];
        const bottomLeft = points[diffs.indexOf(Math.min(...diffs))];
        return [topLeft, topRight, bottomRight, bottomLeft];
    }

    function cropRectFromContour(srcMat, contour) {
        const peri = cv.arcLength(contour, true);
        const approx = new cv.Mat();
        cv.approxPolyDP(contour, approx, 0.04 * peri, true);
        if (approx.rows !== 4) {
            approx.delete();
            return null;
        }

        const points = [];
        for (let i = 0; i < 4; i++) {
            points.push({
                x: approx.intPtr(i, 0)[0],
                y: approx.intPtr(i, 0)[1]
            });
        }
        approx.delete();

        const [tl, tr, br, bl] = orderCornerPoints(points);

        const widthA = Math.hypot(br.x - bl.x, br.y - bl.y);
        const widthB = Math.hypot(tr.x - tl.x, tr.y - tl.y);
        const maxWidth = Math.max(Math.floor(widthA), Math.floor(widthB));

        const heightA = Math.hypot(tr.x - br.x, tr.y - br.y);
        const heightB = Math.hypot(tl.x - bl.x, tl.y - bl.y);
        const maxHeight = Math.max(Math.floor(heightA), Math.floor(heightB));

        if (maxWidth < 80 || maxHeight < 80) return null;

        const srcTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
            tl.x, tl.y, tr.x, tr.y, br.x, br.y, bl.x, bl.y
        ]);
        const dstTri = cv.matFromArray(4, 1, cv.CV_32FC2, [
            0, 0, maxWidth - 1, 0, maxWidth - 1, maxHeight - 1, 0, maxHeight - 1
        ]);
        const M = cv.getPerspectiveTransform(srcTri, dstTri);
        const dst = new cv.Mat();
        cv.warpPerspective(srcMat, dst, M, new cv.Size(maxWidth, maxHeight));
        srcTri.delete();
        dstTri.delete();
        M.delete();

        return dst;
    }

    function computeIoU(a, b) {
        const x1 = Math.max(a.x, b.x);
        const y1 = Math.max(a.y, b.y);
        const x2 = Math.min(a.x + a.w, b.x + b.w);
        const y2 = Math.min(a.y + a.h, b.y + b.h);
        const interW = Math.max(0, x2 - x1);
        const interH = Math.max(0, y2 - y1);
        const interArea = interW * interH;
        if (interArea <= 0) return 0;
        const unionArea = (a.w * a.h) + (b.w * b.h) - interArea;
        return interArea / Math.max(unionArea, 1);
    }

    function preprocessSheet(gray, options) {
        const blur = new cv.Mat();
        cv.GaussianBlur(gray, blur, new cv.Size(5, 5), 0);

        if (!options.enhance) {
            return { base: blur, matsToDelete: [] };
        }

        const equalized = new cv.Mat();
        cv.equalizeHist(blur, equalized);
        blur.delete();

        const denoised = new cv.Mat();
        cv.bilateralFilter(equalized, denoised, 7, 50, 50);
        equalized.delete();

        return { base: denoised, matsToDelete: [] };
    }

    function collectContoursForPass(src, preprocessed, passId, options, candidates, diagnostics) {
        const edges = new cv.Mat();
        const binary = new cv.Mat();
        const contours = new cv.MatVector();
        const hierarchy = new cv.Mat();

        const lower = Math.max(10, 15 + (options.sensitivity * 7));
        const upper = Math.min(230, 90 + (options.sensitivity * 16));

        if (passId === 0) {
            cv.Canny(preprocessed, edges, lower, upper);
            const kernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(3, 3));
            cv.dilate(edges, edges, kernel);
            kernel.delete();
            cv.findContours(edges, contours, hierarchy, cv.RETR_LIST, cv.CHAIN_APPROX_SIMPLE);
        } else if (passId === 1) {
            cv.adaptiveThreshold(preprocessed, binary, 255, cv.ADAPTIVE_THRESH_GAUSSIAN_C, cv.THRESH_BINARY, 21, 5);
            cv.bitwise_not(binary, binary);
            const kernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(4, 4));
            cv.morphologyEx(binary, binary, cv.MORPH_CLOSE, kernel);
            kernel.delete();
            cv.findContours(binary, contours, hierarchy, cv.RETR_EXTERNAL, cv.CHAIN_APPROX_SIMPLE);
        } else {
            cv.threshold(preprocessed, binary, 0, 255, cv.THRESH_BINARY + cv.THRESH_OTSU);
            cv.bitwise_not(binary, binary);
            const kernel = cv.getStructuringElement(cv.MORPH_RECT, new cv.Size(5, 5));
            cv.morphologyEx(binary, binary, cv.MORPH_OPEN, kernel);
            kernel.delete();
            cv.findContours(binary, contours, hierarchy, cv.RETR_LIST, cv.CHAIN_APPROX_SIMPLE);
        }

        const minAreaRatio = 0.002 + ((5 - options.sensitivity) * 0.0005);
        const maxAreaRatio = 0.5;
        const minArea = src.rows * src.cols * minAreaRatio;
        const maxArea = src.rows * src.cols * maxAreaRatio;
        const minDimension = Math.max(70, Math.floor(Math.min(src.cols, src.rows) * 0.07));

        let passAccepted = 0;
        for (let i = 0; i < contours.size(); i++) {
            const contour = contours.get(i);
            const area = cv.contourArea(contour);
            if (area < minArea || area > maxArea) {
                contour.delete();
                continue;
            }

            const rect = cv.boundingRect(contour);
            const contourPerimeter = cv.arcLength(contour, true);
            const compactness = (4 * Math.PI * area) / Math.max(contourPerimeter * contourPerimeter, 1);
            if (rect.width < minDimension || rect.height < minDimension || compactness < 0.15) {
                contour.delete();
                continue;
            }

            const warped = cropRectFromContour(src, contour);
            contour.delete();
            if (!warped) continue;

            const ratio = warped.cols / warped.rows;
            if (ratio < 0.35 || ratio > 2.8) {
                warped.delete();
                continue;
            }

            const bbox = { x: rect.x, y: rect.y, w: rect.width, h: rect.height };
            const duplicate = candidates.some(existing => computeIoU(existing.bbox, bbox) > 0.55);
            if (duplicate) {
                warped.delete();
                continue;
            }

            const qualityGray = new cv.Mat();
            const lap = new cv.Mat();
            cv.cvtColor(warped, qualityGray, cv.COLOR_RGBA2GRAY);
            cv.Laplacian(qualityGray, lap, cv.CV_64F);
            const lapStats = cv.meanStdDev(lap);
            const contrastStats = cv.meanStdDev(qualityGray);
            const sharpness = Math.pow(lapStats.stddev.doubleAt(0, 0), 2);
            const contrast = contrastStats.stddev.doubleAt(0, 0);
            const sizeScore = Math.min((warped.cols * warped.rows) / 220000, 1);
            const qualityScore = (Math.min(sharpness / 220, 1) * 0.5) + (Math.min(contrast / 45, 1) * 0.3) + (sizeScore * 0.2);
            qualityGray.delete();
            lap.delete();
            lapStats.mean.delete();
            lapStats.stddev.delete();
            contrastStats.mean.delete();
            contrastStats.stddev.delete();

            candidates.push({
                mat: warped,
                area: warped.cols * warped.rows,
                bbox,
                passId,
                qualityScore
            });
            passAccepted++;
        }

        diagnostics.push(`Pass ${passId + 1}: ${passAccepted} candidates`);

        edges.delete();
        binary.delete();
        contours.delete();
        hierarchy.delete();
    }

    async function extractPhotosFromSheet(file) {
        if (!isCvReady) {
            throw new Error('Computer vision engine is still loading.');
        }

        const imgEl = await loadImage(file);
        const src = getMatFromImage(imgEl);
        if (!src) throw new Error('Unable to read photo sheet.');

        const scanOptions = {
            sensitivity: parseInt(sheetSensitivityInput?.value || '3', 10),
            maxPhotos: parseInt(sheetMaxPhotosInput?.value || '36', 10),
            enhance: sheetEnhanceInput?.checked !== false
        };
        const gray = new cv.Mat();
        cv.cvtColor(src, gray, cv.COLOR_RGBA2GRAY);
        const prepared = preprocessSheet(gray, scanOptions);
        const diagnostics = [];
        const candidates = [];
        for (let pass = 0; pass < 3; pass++) {
            collectContoursForPass(src, prepared.base, pass, scanOptions, candidates, diagnostics);
            if (candidates.length >= scanOptions.maxPhotos) break;
        }

        candidates.sort((a, b) => (b.qualityScore - a.qualityScore) || (b.area - a.area));
        const limited = candidates.slice(0, scanOptions.maxPhotos);
        candidates.slice(scanOptions.maxPhotos).forEach(item => item.mat.delete());

        const files = [];
        for (let i = 0; i < limited.length; i++) {
            const item = limited[i];
            const canvas = document.createElement('canvas');
            canvas.width = item.mat.cols;
            canvas.height = item.mat.rows;
            cv.imshow(canvas, item.mat);
            item.mat.delete();

            const blob = await new Promise(resolve => canvas.toBlob(resolve, 'image/jpeg', 0.92));
            if (!blob) continue;
            files.push(new File([blob], `sheet_photo_${String(i + 1).padStart(2, '0')}.jpg`, { type: 'image/jpeg' }));
        }

        prepared.base.delete();
        prepared.matsToDelete.forEach(mat => mat.delete());
        src.delete();
        gray.delete();

        return { files, diagnostics };
    }

    // Process Target Images when selected
    async function processTargets(files) {
        if (!isCvReady) {
            progressContainer.classList.remove('hidden');
            progressText.textContent = "Loading Computer Vision engine...";
            return;
        }
        progressContainer.classList.add('hidden');
        
        targetPreview.innerHTML = '';
        
        // Cleanup old target features from C++ memory
        for (let t of targetDescriptorsList) {
            if (t.descriptors && !t.descriptors.isDeleted()) t.descriptors.delete();
            if (t.keypoints && !t.keypoints.isDeleted()) t.keypoints.delete();
            if (t.hiResFeatures) {
                if (t.hiResFeatures.descriptors && !t.hiResFeatures.descriptors.isDeleted()) t.hiResFeatures.descriptors.delete();
                if (t.hiResFeatures.keypoints && !t.hiResFeatures.keypoints.isDeleted()) t.hiResFeatures.keypoints.delete();
                t.hiResFeatures = null;
            }
        }
        targetDescriptorsList = [];

        for (const file of files) {
            const imgEl = await loadImage(file);
            imgEl.title = file.name;
            // style preview
            imgEl.style.width = '60px';
            imgEl.style.height = '60px';
            imgEl.style.objectFit = 'cover';
            imgEl.style.borderRadius = '8px';
            targetPreview.appendChild(imgEl);

            try {
                const feats = extractFeatures(imgEl, { maxDim: 600, orbFeatures: 500 });
                if (feats && feats.descriptors.rows > 0) {
                    targetDescriptorsList.push({ file, ...feats });
                }
            } catch (err) {
                console.error('Error extracting features target', file.name, err);
            }
        }
        checkReadyState();
    }

    // Upload Listeners
    targetInput.addEventListener('change', async (e) => {
        const files = Array.from(e.target.files);
        if (files.length > 0) {
            await processTargets(files);
        }
    });

    if (sheetImageInput) {
        sheetImageInput.addEventListener('change', async (e) => {
            pendingSheetImage = e.target.files[0] || null;
            scanSheetBtn.classList.toggle('hidden', !pendingSheetImage);
            sheetStatus.textContent = pendingSheetImage
                ? `Loaded "${pendingSheetImage.name}". Click Auto-Detect Photos.`
                : 'Tip: works best with top-down photos and good lighting.';
        });
    }

    if (scanSheetBtn) {
        scanSheetBtn.addEventListener('click', async () => {
            if (!pendingSheetImage) return;
            scanSheetBtn.disabled = true;
            scanSheetBtn.textContent = 'Detecting...';
            sheetStatus.textContent = 'Scanning photo sheet and extracting each printed photo...';

            try {
                const scanResult = await extractPhotosFromSheet(pendingSheetImage);
                const extractedFiles = scanResult.files;
                if (extractedFiles.length === 0) {
                    sheetStatus.textContent = 'No photo blocks detected. Try a sharper image with less glare.';
                } else {
                    await processTargets(extractedFiles);
                    sheetStatus.textContent = `Detected ${extractedFiles.length} photos and loaded them as targets. ${scanResult.diagnostics.join(' • ')}`;
                }
            } catch (err) {
                console.error('Photo sheet scan failed', err);
                sheetStatus.textContent = err.message || 'Could not scan this sheet image.';
            } finally {
                scanSheetBtn.disabled = false;
                scanSheetBtn.textContent = 'Auto-Detect Photos';
            }
        });
    }

    directoryInput.addEventListener('change', (e) => {
        const files = Array.from(e.target.files);
        clearFeatureCache(); // reset cache when directory scope changes
        
        searchFiles = files.filter(file => 
            file.type.startsWith('image/') || 
            file.name.match(/\.(jpg|jpeg|png|gif|webp|bmp|tif|tiff)$/i)
        );
        
        if (searchFiles.length > 0) {
            directoryInfo.textContent = `Found ${searchFiles.length} images to scan.`;
            directoryInfo.style.color = "var(--success)";
        } else {
            directoryInfo.textContent = 'No images found in the selected directory.';
            directoryInfo.style.color = "#ef4444";
        }
        checkReadyState();
    });

    // Check button status
    function checkReadyState() {
        if (isCvReady && targetDescriptorsList.length > 0 && searchFiles.length > 0) {
            searchBtn.disabled = false;
            searchBtn.textContent = "Search with Computer Vision";
        } else if (!isCvReady) {
            searchBtn.disabled = true;
            searchBtn.textContent = "Loading AI Engine...";
        } else {
            searchBtn.disabled = true;
            searchBtn.textContent = "Search for Matches";
        }
    }
    
    // Safety init check
    checkReadyState();

    // Helper: Compare descriptors with Lowe's ratio + geometric verification (RANSAC)
    function evaluateMatch(targetFeatures, queryFeatures) {
        const desc1 = targetFeatures.descriptors;
        const desc2 = queryFeatures.descriptors;
        if (desc1.rows === 0 || desc2.rows === 0) {
            return { isMatch: false, confidence: 'Low', score: 0, goodMatches: 0, inliers: 0, inlierRatio: 0 };
        }
        
        // NORM_HAMMING is necessary for ORB (binary descriptors)
        const bf = new cv.BFMatcher(cv.NORM_HAMMING, false);
        const matches = new cv.DMatchVectorVector();
        let isMatched = false;
        let confidence = 'Low';
        let score = 0;
        let inliers = 0;
        let inlierRatio = 0;
        
        try {
            bf.knnMatch(desc2, desc1, matches, 2);
        } catch (e) {
            console.error("KNN Match failed", e);
            bf.delete();
            matches.delete();
            return { isMatch: false, confidence: 'Low', score: 0, goodMatches: 0, inliers: 0, inlierRatio: 0 };
        }

        let goodMatches = 0;
        const srcPoints = []; // target points
        const dstPoints = []; // query/search points
        // Lowe's ratio test filter
        const RATIO_THRESHOLD = 0.75;
        
        for (let i = 0; i < matches.size(); ++i) {
            let match = matches.get(i);
            if (match.size() >= 2) {
                const dMatch1 = match.get(0);
                const dMatch2 = match.get(1);
                if (dMatch1.distance <= RATIO_THRESHOLD * dMatch2.distance) {
                    goodMatches++;
                    const targetKp = targetFeatures.keypoints.get(dMatch1.trainIdx);
                    const queryKp = queryFeatures.keypoints.get(dMatch1.queryIdx);
                    srcPoints.push(targetKp.pt.x, targetKp.pt.y);
                    dstPoints.push(queryKp.pt.x, queryKp.pt.y);
                }
            } else if (match.size() === 1) {
                const dMatch = match.get(0);
                goodMatches++;
                const targetKp = targetFeatures.keypoints.get(dMatch.trainIdx);
                const queryKp = queryFeatures.keypoints.get(dMatch.queryIdx);
                srcPoints.push(targetKp.pt.x, targetKp.pt.y);
                dstPoints.push(queryKp.pt.x, queryKp.pt.y);
            }
        }

        // Adaptive minimum good matches scales by descriptor count
        const minGoodMatches = Math.max(8, Math.floor(0.08 * Math.min(desc1.rows, desc2.rows)));

        // Geometric verification using homography and RANSAC
        if (goodMatches >= 4 && srcPoints.length >= 8 && dstPoints.length >= 8) {
            const srcMat = cv.matFromArray(goodMatches, 1, cv.CV_32FC2, srcPoints);
            const dstMat = cv.matFromArray(goodMatches, 1, cv.CV_32FC2, dstPoints);
            const inlierMask = new cv.Mat();
            let H = null;

            try {
                H = cv.findHomography(srcMat, dstMat, cv.RANSAC, 3.0, inlierMask);
                if (H && !H.empty()) {
                    for (let i = 0; i < inlierMask.rows; i++) {
                        if (inlierMask.data[i]) inliers++;
                    }
                    inlierRatio = inliers / Math.max(goodMatches, 1);
                }
            } catch (e) {
                console.error("Homography check failed", e);
            } finally {
                if (H) H.delete();
                srcMat.delete();
                dstMat.delete();
                inlierMask.delete();
            }
        }

        const minInliers = Math.max(6, Math.floor(minGoodMatches * 0.5));
        isMatched = goodMatches >= minGoodMatches && inliers >= minInliers && inlierRatio >= 0.35;

        // Confidence + score for UI/debugging
        score = Math.round((Math.min(goodMatches / 40, 1) * 35) + (Math.min(inliers / 25, 1) * 35) + (inlierRatio * 30));
        if (isMatched && inlierRatio >= 0.6 && inliers >= 20) confidence = 'High';
        else if (isMatched && inlierRatio >= 0.45 && inliers >= 10) confidence = 'Medium';
        else confidence = 'Low';

        bf.delete();
        matches.delete();

        return { isMatch: isMatched, confidence, score, goodMatches, inliers, inlierRatio };
    }

    // Execute Search Logic
    searchBtn.addEventListener('click', async () => {
        // Reveal UI
        resultsSection.classList.remove('hidden');
        progressContainer.classList.remove('hidden');
        clearMemoryBtn.classList.remove('hidden');
        searchBtn.disabled = true;
        
        let matchesFound = 0;
        let scanned = 0;

        // Visual distinction for multiple searches
        if (resultsGrid.children.length > 0) {
            const divider = document.createElement('div');
            divider.className = 'search-divider';
            const timeStr = new Date().toLocaleTimeString([], {hour: '2-digit', minute:'2-digit'});
            divider.innerHTML = `<div>New Search Session</div> <span>${timeStr} &bull; Targets: ${targetDescriptorsList.length} comparing ${searchFiles.length} files</span>`;
            resultsGrid.appendChild(divider);
        }

        for (const file of searchFiles) {
            scanned++;
            progressText.textContent = `Analyzing ${scanned} of ${searchFiles.length}...`;
            
            // Allow UI to repaint every 2 files
            if (scanned % 2 === 0) {
                await yieldToMain();
            }

            try {
                const fastFeats = await getOrCreateFileFeatures(file, { maxDim: 600, orbFeatures: 500 });
                
                if (fastFeats && fastFeats.descriptors.rows > 0) {
                    // Check against all target images and keep best-scoring verified match
                    let bestMatch = null;
                    for (const target of targetDescriptorsList) {
                        const matchResult = evaluateMatch(target, fastFeats);
                        if (matchResult.isMatch) {
                            if (!bestMatch || matchResult.score > bestMatch.matchResult.score) {
                                bestMatch = { target, matchResult };
                            }
                        }
                    }

                    // High-resolution recovery pass if first pass did not confidently match
                    if (!bestMatch) {
                        const hiResQuery = await getOrCreateFileFeatures(file, { maxDim: 1024, orbFeatures: 1200 });
                        if (hiResQuery && hiResQuery.descriptors.rows > 0) {
                            for (const target of targetDescriptorsList) {
                                const hiResTarget = await getOrCreateHiResTargetFeatures(target);
                                if (!hiResTarget) continue;
                                const matchResult = evaluateMatch(hiResTarget, hiResQuery);
                                if (matchResult.isMatch) {
                                    matchResult.stage = 'High-Res verification';
                                    if (!bestMatch || matchResult.score > bestMatch.matchResult.score) {
                                        bestMatch = { target, matchResult };
                                    }
                                }
                            }
                        }
                    }

                    if (bestMatch) {
                        matchesFound++;
                        displayMatch(file, bestMatch.target.file.name, bestMatch.matchResult);
                    }
                }
            } catch (err) {
                console.error('Failed to process search file:', file.name, err);
            }
        }

        progressContainer.classList.add('hidden');
        if (matchesFound === 0) {
            const noRes = document.createElement('div');
            noRes.className = 'no-results';
            noRes.textContent = 'No similar images (cropped, rotated, or parts) found in this session.';
            resultsGrid.appendChild(noRes);
        }
        searchBtn.disabled = false;
    });

    // Provide ability to wipe all memory
    clearMemoryBtn.addEventListener('click', async () => {
        if (!confirm('Are you sure you want to clear all searched results and wipe local database memory?')) return;
        selectedFiles = [];
        resultsGrid.innerHTML = '';
        exportBar.classList.add('hidden');
        resultsSection.classList.add('hidden');
        clearMemoryBtn.classList.add('hidden');
        clearFeatureCache();
        
        await clearDB();
    });

    // --- Save/Load Progress System ---
    if (saveProgressBtn) {
        saveProgressBtn.addEventListener('click', async () => {
            if (selectedFiles.length === 0) {
                alert("No images selected! Please select some images before saving progress.");
                return;
            }
            
            saveProgressBtn.disabled = true;
            const originalText = saveProgressBtn.textContent;
            saveProgressBtn.textContent = 'Saving...';
            
            try {
                const zip = new JSZip();
                for (let i = 0; i < selectedFiles.length; i++) {
                    const f = selectedFiles[i].file;
                    const buffer = await f.arrayBuffer();
                    // Keep order chronological with zero pad
                    const prepName = String(i).padStart(4, '0') + "_" + f.name;
                    zip.file(prepName, buffer);
                }
                
                const content = await zip.generateAsync({ type: "blob" });
                const url = URL.createObjectURL(content);
                const a = document.createElement("a");
                a.style.display = 'none';
                a.href = url;
                a.download = "My_Image_Project.zip";
                document.body.appendChild(a);
                a.click();
                URL.revokeObjectURL(url);
                document.body.removeChild(a);
            } catch(e) {
                console.error("Save Error:", e);
                alert("Failed to save progress.");
            } finally {
                saveProgressBtn.disabled = false;
                saveProgressBtn.textContent = originalText;
            }
        });

        loadProgressBtn.addEventListener('click', () => loadProgressInput.click());

        loadProgressInput.addEventListener('change', async (e) => {
            const file = e.target.files[0];
            if (!file) return;
            
            loadProgressBtn.disabled = true;
            const originalText = loadProgressBtn.textContent;
            loadProgressBtn.textContent = 'Restoring...';
            
            try {
                const zip = await JSZip.loadAsync(file);
                const sortedFilenames = Object.keys(zip.files).sort();
                
                resultsSection.classList.remove('hidden');
                clearMemoryBtn.classList.remove('hidden');
                
                const divider = document.createElement('div');
                divider.className = 'search-divider';
                divider.innerHTML = `<div>Project Restored</div> <span>Loaded exactly from sequence</span>`;
                if (resultsGrid) resultsGrid.appendChild(divider);
                
                for (const filename of sortedFilenames) {
                    if (zip.files[filename].dir) continue;
                    
                    const blob = await zip.files[filename].async("blob");
                    const originalName = filename.substring(filename.indexOf('_') + 1);
                    
                    const newFile = new File([blob], originalName, { type: blob.type || "image/jpeg" });
                    Object.defineProperty(newFile, 'webkitRelativePath', { value: originalName });
                    
                    const card = displayMatch(newFile, "Restored Progress");
                    card.classList.add('selected');
                    const badge = card.querySelector('.selection-badge');
                    selectedFiles.push({ file: newFile, card, badge });
                    
                    // Hook into underlying IndexedDB defensively
                    saveToDB(newFile);
                }
                
                updateExportBar();
                alert("Progress Restored successfully! You can jump straight to Spreads.");
                
            } catch(e) {
                console.error("Load Error:", e);
                alert("Failed to read progress file. Ensure it is a valid project ZIP.");
            } finally {
                loadProgressBtn.disabled = false;
                loadProgressBtn.textContent = originalText;
                loadProgressInput.value = '';
            }
        });
    }

    function displayMatch(file, matchedTargetName, matchMeta = null) {
        const url = URL.createObjectURL(file);
        
        const card = document.createElement('div');
        card.className = 'result-card';
        card.style.animationDelay = `${(resultsGrid.children.length % 10) * 0.1}s`;
        
        const img = document.createElement('img');
        img.className = 'result-image';
        img.src = url;
        img.loading = "lazy";
        
        const details = document.createElement('div');
        details.className = 'result-details';
        
        const name = document.createElement('div');
        name.className = 'result-filename';
        name.textContent = file.name;
        
        const path = document.createElement('div');
        path.className = 'result-path';
        path.textContent = file.webkitRelativePath || file.name;
        
        const matchInfo = document.createElement('div');
        matchInfo.className = 'result-path';
        matchInfo.style.color = 'var(--primary)';
        matchInfo.style.marginTop = '0.4rem';
        let matchText = `Matched: ${matchedTargetName}`;
        if (matchMeta) {
            const ratioPct = Math.round((matchMeta.inlierRatio || 0) * 100);
            matchText += ` • ${matchMeta.confidence} confidence`;
            matchText += ` (score ${matchMeta.score}, inliers ${matchMeta.inliers}/${matchMeta.goodMatches}, ${ratioPct}% inlier ratio)`;
            if (matchMeta.stage) matchText += ` • ${matchMeta.stage}`;
        }
        matchInfo.textContent = matchText;

        details.appendChild(name);
        details.appendChild(path);
        details.appendChild(matchInfo);
        card.appendChild(img);
        card.appendChild(details);
        
        const badge = document.createElement('div');
        badge.className = 'selection-badge';
        card.appendChild(badge);

        // Toggle selection with array ordering
        card.addEventListener('click', () => {
            const idx = selectedFiles.findIndex(f => f.file === file);
            if (idx > -1) {
                card.classList.remove('selected');
                selectedFiles.splice(idx, 1);
                removeFromDB(file.name);
            } else {
                card.classList.add('selected');
                selectedFiles.push({ file, card, badge });
                saveToDB(file);
            }
            updateExportBar();
        });
        
        resultsGrid.appendChild(card);
        return card;
    }
    
    // UI Update for Export Bar
    function updateExportBar() {
        // Update numbering
        selectedFiles.forEach((item, index) => {
            item.badge.textContent = index + 1;
        });

        if (selectedFiles.length > 0) {
            exportBar.classList.remove('hidden');
            selectionCount.textContent = `${selectedFiles.length} selected`;
        } else {
            exportBar.classList.add('hidden');
        }
    }


    /* ========================================================================= */
    /* ==================== DRAG AND DROP SPREADS WORKSPACE ==================== */
    /* ========================================================================= */

    // Helper: Make an element a dropzone
    function setupDropzone(zone) {
        zone.addEventListener('dragover', (e) => {
            e.preventDefault(); // allow dropping
            e.dataTransfer.dropEffect = 'move';
            zone.classList.add('drag-over');
        });
        
        zone.addEventListener('dragleave', () => {
            zone.classList.remove('drag-over');
        });
        
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('drag-over');
            
            if (draggedElement) {
                // If dropping into a page that already has an image, swap them
                if (zone.classList.contains('spread-page') && zone.children.length > 0) {
                    const existingImg = zone.children[0];
                    const sourceParent = draggedElement.parentNode;
                    // Move existing image to the source zone
                    sourceParent.appendChild(existingImg);
                    // Move dragged image to this zone
                    zone.appendChild(draggedElement);
                } else {
                    // Empty zone or Images Dock allows infinite appends
                    zone.appendChild(draggedElement);
                }
            }
        });
    }

    // Helper: Create a single blank 2-page spread
    function createSpreadDOM() {
        const spread = document.createElement('div');
        spread.className = 'spread-layout';
        
        const leftPage = document.createElement('div');
        leftPage.className = 'spread-page dropzone';
        
        const rightPage = document.createElement('div');
        rightPage.className = 'spread-page dropzone';
        
        spread.appendChild(leftPage);
        spread.appendChild(rightPage);
        spreadsContainer.appendChild(spread);
        
        setupDropzone(leftPage);
        setupDropzone(rightPage);
        return { leftPage, rightPage };
    }

    // Initialize Workspace dock dropzone
    setupDropzone(imagesDock);

    // Open Workspace
    organizeSpreadsBtn.addEventListener('click', () => {
        if (selectedFiles.length === 0) return;
        
        // Reset DOM state
        spreadsContainer.innerHTML = '';
        imagesDock.innerHTML = '';
        
        // Auto-generate enough spreads to hold our selection (2 images per spread)
        const numSpreads = Math.ceil(selectedFiles.length / 2);
        const autoPages = [];
        for(let i=0; i<numSpreads; i++) {
            const { leftPage, rightPage } = createSpreadDOM();
            autoPages.push(leftPage, rightPage);
        }
        
        // Create draggable image elements and auto-flow them
        selectedFiles.forEach((fileObj, index) => {
            const img = document.createElement('img');
            img.src = URL.createObjectURL(fileObj.file);
            img.className = 'draggable-image';
            img.draggable = true;
            // Map the internal file array index to the DOM object so PDF generator can access it
            img.dataset.fileIndex = index; 
            
            // Drag Events
            img.addEventListener('dragstart', (e) => {
                draggedElement = img;
                img.classList.add('dragging');
                e.dataTransfer.effectAllowed = 'move';
            });
            img.addEventListener('dragend', () => {
                img.classList.remove('dragging');
                draggedElement = null;
            });
            
            // Auto-flow sequentially into pages, else drop to Dock if array out of bounds
            if (index < autoPages.length) {
                autoPages[index].appendChild(img);
            } else {
                imagesDock.appendChild(img);
            }
        });
        
        // Show Workspace Overlay
        workspace.classList.remove('hidden');
    });

    closeWorkspaceBtn.addEventListener('click', () => {
        workspace.classList.add('hidden');
    });
    
    addSpreadBtn.addEventListener('click', () => {
        createSpreadDOM();
    });

    // Spreads Workspace visual drag-and-drop export logic
    workspaceExportPdfBtn.addEventListener('click', async () => {
        const spreads = document.querySelectorAll('.spread-layout');
        if (spreads.length === 0) return;
        
        workspaceExportPdfBtn.disabled = true;
        workspaceExportPdfBtn.textContent = 'Generating PDF...';

        try {
            const { jsPDF } = window.jspdf;
            const doc = new jsPDF({ orientation: 'landscape', format: 'a4' });
            
            const margin = 15;
            const pageWidth = doc.internal.pageSize.getWidth();
            const pageHeight = doc.internal.pageSize.getHeight();
            const usableWidth = pageWidth - (margin * 3); 
            const spreadWidth = usableWidth / 2;
            const usableHeight = pageHeight - margin * 2 - 20;

            // Helper to draw image into bounded box on PDF
            async function drawImageToPdf(imgElement, startX, startY) {
                if (!imgElement) return;
                
                const fIdx = parseInt(imgElement.dataset.fileIndex);
                if(isNaN(fIdx)) return;
                
                const fileObj = selectedFiles[fIdx];
                const img = await loadImage(fileObj.file);
                const canvas = document.createElement('canvas');
                canvas.width = img.width;
                canvas.height = img.height;
                const ctx = canvas.getContext('2d');
                ctx.drawImage(img, 0, 0);
                const jpegData = canvas.toDataURL('image/jpeg', 0.9);
                
                let w = img.width;
                let h = img.height;
                
                if (w > spreadWidth || h > usableHeight) {
                    const ratio = Math.min(spreadWidth / w, usableHeight / h);
                    w *= ratio;
                    h *= ratio;
                }
                
                const xOffset = startX + (spreadWidth - w) / 2;
                const yOffset = startY + (usableHeight - h) / 2;
                
                doc.addImage(jpegData, 'JPEG', xOffset, yOffset, w, h);
            }

            let pageIndex = 0;
            // Iterate over the visible DOM spreads
            for (let i = 0; i < spreads.length; i++) {
                const spread = spreads[i];
                // Access the left `.page` child and right `.page` child's IMG tag
                const leftImg = spread.children[0].querySelector('img');
                const rightImg = spread.children[1].querySelector('img');
                
                // Skip completely blank spreads
                if (!leftImg && !rightImg) continue;

                if (pageIndex > 0) doc.addPage();
                
                // Left Spread Image
                if (leftImg) await drawImageToPdf(leftImg, margin, margin + 10);
                // Right Spread Image
                if (rightImg) await drawImageToPdf(rightImg, margin * 2 + spreadWidth, margin + 10);
                
                pageIndex++;
            }
            
            doc.save('Custom_Spreads.pdf');
            
            document.getElementById('modal-title').textContent = '🎉 Book Spreads Exported!';
            document.getElementById('modal-instructions').innerHTML = `
                <h4>How to insert this into Adobe InDesign:</h4>
                <ol>
                    <li>Open your Adobe InDesign document.</li>
                    <li>Press <kbd>Ctrl</kbd> + <kbd>D</kbd> (Windows) or <kbd>Cmd</kbd> + <kbd>D</kbd> (Mac) to open the <strong>Place</strong> dialog.</li>
                    <li>Select the downloaded <code>Custom_Spreads.pdf</code> file.</li>
                    <li>Because you designed spreads, you can select the "Import Options" and place these wide 2-page chunks directly into your matching templates!</li>
                </ol>
            `;
            successModal.classList.remove('hidden');
            
        } catch (e) {
            console.error('PDF generation failed', e);
            alert('Failed to generate PDF. Make sure you are connected to the internet.');
        } finally {
            workspaceExportPdfBtn.disabled = false;
            workspaceExportPdfBtn.textContent = 'Export Spreads to PDF';
        }
    });

    /* ========================================================================= */
    /* ========================= INDESIGN EXPORT =============================== */
    /* ========================================================================= */
    
    const workspaceExportIndesignBtn = document.getElementById('workspace-export-indesign-btn');
    if (workspaceExportIndesignBtn) {
        workspaceExportIndesignBtn.addEventListener('click', async () => {
            const spreads = document.querySelectorAll('.spread-layout');
            if (spreads.length === 0) return;
            
            workspaceExportIndesignBtn.disabled = true;
            const oText = workspaceExportIndesignBtn.textContent;
            workspaceExportIndesignBtn.textContent = 'Building InDesign Project...';

            try {
                const zip = new JSZip();
                const imgFolder = zip.folder("images");
                
                let spreadsData = [];
                
                // Loop through spreads to copy images and build script data
                for (let i = 0; i < spreads.length; i++) {
                    const spread = spreads[i];
                    const leftImg = spread.children[0].querySelector('img');
                    const rightImg = spread.children[1].querySelector('img');
                    
                    if (!leftImg && !rightImg) continue;
                    
                    let spreadObj = { left: null, right: null };
                    
                    if (leftImg) {
                        const fIdx = parseInt(leftImg.dataset.fileIndex);
                        const fileObj = selectedFiles[fIdx];
                        const originalName = fileObj.file.name;
                        spreadObj.left = originalName;
                        
                        const buffer = await fileObj.file.arrayBuffer();
                        imgFolder.file(originalName, buffer);
                    }
                    
                    if (rightImg) {
                        const fIdx = parseInt(rightImg.dataset.fileIndex);
                        const fileObj = selectedFiles[fIdx];
                        const originalName = fileObj.file.name;
                        spreadObj.right = originalName;
                        
                        const buffer = await fileObj.file.arrayBuffer();
                        imgFolder.file(originalName, buffer);
                    }
                    
                    spreadsData.push(spreadObj);
                }
                
                // Generate the InDesign ExtendScript (.jsx) file content
                const jsxCode = `// Auto-generated InDesign Build Script
#target indesign

var scriptFile = new File($.fileName);
var scriptFolder = scriptFile.parent;
var imagesFolder = new Folder(scriptFolder.fsName + "/images");

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;

var doc = app.documents.add();
doc.documentPreferences.facingPages = false;
doc.documentPreferences.pageWidth = "420mm"; // Wide document holding two A4 sections
doc.documentPreferences.pageHeight = "297mm";

var spreadsData = ${JSON.stringify(spreadsData)};

for (var i = 0; i < spreadsData.length; i++) {
    var page;
    if (i === 0) {
        page = doc.pages[0];
    } else {
        page = doc.pages.add();
    }
    
    var spreadDef = spreadsData[i];
    
    // Left Page Setup
    if (spreadDef.left) {
        var imgFile = new File(imagesFolder.fsName + "/" + spreadDef.left);
        if (imgFile.exists) {
            var rect = page.rectangles.add();
            // Bounds layout: [Y1, X1, Y2, X2]
            rect.geometricBounds = ["0mm", "0mm", "297mm", "210mm"];
            rect.place(imgFile);
            rect.fit(FitOptions.PROPORTIONALLY);
            rect.fit(FitOptions.CENTER_CONTENT);
        }
    }
    
    // Right Page Setup
    if (spreadDef.right) {
        var imgFile = new File(imagesFolder.fsName + "/" + spreadDef.right);
        if (imgFile.exists) {
            var rect = page.rectangles.add();
            // Bounds layout: [Y1, X1, Y2, X2]
            rect.geometricBounds = ["0mm", "210mm", "297mm", "420mm"];
            rect.place(imgFile);
            rect.fit(FitOptions.PROPORTIONALLY);
            rect.fit(FitOptions.CENTER_CONTENT);
        }
    }
}

app.scriptPreferences.userInteractionLevel = UserInteractionLevels.INTERACT_WITH_ALL;
alert("Your Book Spreads were successfully matched and built!");
`;

                zip.file("1_Click_Build_InDesign_Spreads.jsx", jsxCode);
                
                // Download the ZIP payload
                const content = await zip.generateAsync({ type: "blob" });
                const url = URL.createObjectURL(content);
                const a = document.createElement("a");
                a.style.display = 'none';
                a.href = url;
                a.download = "InDesign_Spreads_Project.zip";
                document.body.appendChild(a);
                a.click();
                URL.revokeObjectURL(url);
                document.body.removeChild(a);
                
                document.getElementById('modal-title').textContent = '🎉 InDesign Project Generated!';
                document.getElementById('modal-instructions').innerHTML = `
                    <h4>How to instantly assemble your InDesign Document:</h4>
                    <ol>
                        <li>Extract the downloaded <code>InDesign_Spreads_Project.zip</code> folder anywhere on your computer.</li>
                        <li>Inside, you will find an <code>images</code> folder and a script named <code>1_Click_Build_InDesign_Spreads.jsx</code>.</li>
                        <li>Open Adobe InDesign.</li>
                        <li><b>Drag and drop the <code>.jsx</code> file directly into the InDesign window</b> (or double-click it).</li>
                        <li>InDesign will automatically run the script, instantly create a brand new layout, and perfectly place all your photos exactly as you ordered them!</li>
                    </ol>
                `;
                successModal.classList.remove('hidden');

            } catch (e) {
                console.error('InDesign export failed', e);
                alert('Failed to generate InDesign Project. Ensure internet access for JSZip.');
            } finally {
                workspaceExportIndesignBtn.disabled = false;
                workspaceExportIndesignBtn.textContent = oText;
            }
        });
    }

    /* ========================================================================= */
    /* ============================== ZIP EXPORT =============================== */
    /* ========================================================================= */

    exportFolderBtn.addEventListener('click', async () => {
        if (selectedFiles.length === 0) return;
        
        exportFolderBtn.disabled = true;
        const originalText = exportFolderBtn.textContent;
        exportFolderBtn.textContent = 'Zipping...';

        try {
            const zip = new JSZip();
            const imgFolder = zip.folder("Matched_Images");

            for (let i = 0; i < selectedFiles.length; i++) {
                const fileObj = selectedFiles[i];
                // Read as ArrayBuffer for JSZip
                const buffer = await fileObj.file.arrayBuffer();
                // Keep original filename as requested previously
                imgFolder.file(fileObj.file.name, buffer);
            }

            const content = await zip.generateAsync({ type: "blob" });
            const url = URL.createObjectURL(content);
            const a = document.createElement("a");
            a.style.display = 'none';
            a.href = url;
            a.download = "Matched_Images.zip";
            document.body.appendChild(a);
            a.click();
            URL.revokeObjectURL(url);
            document.body.removeChild(a);
            
            document.getElementById('modal-title').textContent = '🎉 Folder Exported Successfully!';
            document.getElementById('modal-instructions').innerHTML = `
                <h4>How to use your new folder:</h4>
                <ol>
                    <li>Extract the downloaded <code>Matched_Images.zip</code> file (right-click and select "Extract All").</li>
                    <li>Inside, you will find a folder containing all your exact images with their original filenames.</li>
                    <li>You can now drag and drop them directly into whatever application you choose!</li>
                </ol>
            `;
            successModal.classList.remove('hidden');

        } catch (e) {
            console.error('ZIP generation failed', e);
            alert('Failed to generate ZIP folder. Make sure you have internet access to load JSZip.');
        } finally {
            exportFolderBtn.disabled = false;
            exportFolderBtn.textContent = originalText;
        }
    });
});
