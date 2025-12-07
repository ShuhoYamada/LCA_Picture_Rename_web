/**
 * Image Renamer Pro - Web Version
 * Logic for handling Excel imports, image processing, and file renaming.
 */

// --- State Management ---
const state = {
    materials: {},          // Display Name -> ID
    materialsData: {},      // ID -> Details
    materialCategories: {}, // Category -> [Material Names]
    materialNameToId: {},   // Material Name -> ID

    processing: {},         // Display Name -> ID
    processingData: {},     // ID -> Details

    images: [],             // Array of File objects
    currentIndex: 0,        // Current image index

    renamedFiles: {},       // Original Filename -> New Filename (Map)

    // Per-image persistence (Part Name & Weight)
    persistedData: {
        partName: "",
        weight: "",
        unit: "kg",
        materialCategory: "",
        material: "",
        processing: "",
        photoType: "P",
        notes: "0"
    }
};

// --- DOM Elements ---
const dom = {
    app: document.getElementById('app'),

    // File Inputs
    fileMaterial: document.getElementById('file-material'),
    fileProcess: document.getElementById('file-process'),
    fileImages: document.getElementById('file-images'),

    // Buttons
    btnApply: document.getElementById('btn-apply'),
    btnDownload: document.getElementById('btn-download'),
    btnReset: document.getElementById('btn-reset'),
    btnPrev: document.getElementById('nav-prev'),
    btnNext: document.getElementById('nav-next'),

    // Status Indicators
    statusMaterial: document.getElementById('status-material'),
    statusProcess: document.getElementById('status-process'),
    statusImages: document.getElementById('status-images'),

    // Inputs
    inputPartName: document.getElementById('input-part-name'),
    inputWeight: document.getElementById('input-weight'),
    inputUnit: document.getElementById('input-unit'),
    inputCategory: document.getElementById('input-material-category'),
    inputMaterial: document.getElementById('input-material'),
    inputProcessing: document.getElementById('input-processing'),
    inputPhotoType: document.getElementById('input-photo-type'),
    inputNotes: document.getElementById('input-notes'),

    // Viewer
    previewImage: document.getElementById('preview-image'),
    placeholder: document.getElementById('placeholder-text'),
    imageCounter: document.getElementById('image-counter'),
    currentFilename: document.getElementById('current-filename'),
    renamePreview: document.getElementById('rename-preview')
};

// --- Initialization ---
function init() {
    setupEventListeners();
    updateUIState();
}

function setupEventListeners() {
    // File Inputs
    dom.fileMaterial.addEventListener('change', handleMaterialLoad);
    dom.fileProcess.addEventListener('change', handleProcessingLoad);
    dom.fileImages.addEventListener('change', handleImagesLoad);

    // Navigation
    dom.btnPrev.addEventListener('click', () => navigateImage(-1));
    dom.btnNext.addEventListener('click', () => navigateImage(1));

    // Actions
    dom.btnApply.addEventListener('click', applyAndNext);
    dom.btnDownload.addEventListener('click', downloadZip);
    dom.btnReset.addEventListener('click', resetApp);

    // Input Changes (for validation & preview)
    const inputs = [
        dom.inputPartName, dom.inputWeight, dom.inputUnit,
        dom.inputCategory, dom.inputMaterial, dom.inputProcessing,
        dom.inputPhotoType, dom.inputNotes
    ];
    inputs.forEach(input => {
        input.addEventListener('input', () => {
            validateInputs();
            updateRenamePreview();
        });
        input.addEventListener('change', () => {
            validateInputs();
            updateRenamePreview();
        });
    });

    // Material Category Change
    dom.inputCategory.addEventListener('change', handleCategoryChange);

    // Keyboard Shortcuts
    document.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !dom.btnApply.disabled) {
            applyAndNext();
        } else if (e.key === 'ArrowLeft') {
            navigateImage(-1);
        } else if (e.key === 'ArrowRight') {
            // Only navigating, not applying
            navigateImage(1);
        }
    });

    // Weight Validation (Allow text input but validate content)
    dom.inputWeight.addEventListener('input', (e) => {
        const val = e.target.value;
        if (val && !/^[0-9a-zA-Z.-]*$/.test(val)) {
            e.target.classList.add('input-error');
        } else {
            e.target.classList.remove('input-error');
        }
    });
}

// --- Excel Handling ---

async function handleMaterialLoad(e) {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const rows = await readXlsxFile(file);
        parseMaterialMaster(rows);

        updateStatus(dom.statusMaterial, true, '素材マスタ読込済');
        dom.fileMaterial.parentElement.classList.replace('btn-secondary', 'btn-success');
        updateUIState();
    } catch (err) {
        console.error(err);
        alert('素材マスタの読み込みに失敗しました。');
    }
}

function parseMaterialMaster(rows) {
    if (rows.length < 2) return;

    // Header analysis
    const header = rows[0];

    // Allow '素材名' or '素材'
    let colName = header.indexOf('素材名');
    if (colName === -1) {
        colName = header.indexOf('素材');
    }

    const colId = header.indexOf('素材ID');

    // Category is optional now to handle simple lists
    let colCat = header.indexOf('素材区分');

    if (colName === -1 || colId === -1) {
        alert('エラー：素材マスタに必要な列（素材名, 素材ID）が見つかりません。');
        return;
    }

    // Reset state
    state.materials = {};
    state.materialCategories = {};
    state.materialNameToId = {};

    // Parse rows
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const name = row[colName];
        const id = row[colId];
        // If category column exists, use it. Otherwise use "共通" (Common)
        const cat = colCat !== -1 ? row[colCat] : "共通";

        if (name && id) {
            const catStr = cat || "共通"; // Safety for empty cells in category column

            // Category mapping
            if (!state.materialCategories[catStr]) {
                state.materialCategories[catStr] = [];
            }
            state.materialCategories[catStr].push(name);

            // ID mapping
            state.materialNameToId[name] = id;
            state.materials[`${name}`] = id;
        }
    }

    // Update Category Dropdown
    dom.inputCategory.innerHTML = '<option value="">選択してください</option>';
    Object.keys(state.materialCategories).forEach(cat => {
        const option = document.createElement('option');
        option.value = cat;
        option.textContent = cat;
        dom.inputCategory.appendChild(option);
    });
}

async function handleProcessingLoad(e) {
    const file = e.target.files[0];
    if (!file) return;

    try {
        const rows = await readXlsxFile(file);
        parseProcessingMaster(rows);

        updateStatus(dom.statusProcess, true, '加工マスタ読込済');
        dom.fileProcess.parentElement.classList.replace('btn-warning', 'btn-success');
        updateUIState();
    } catch (err) {
        console.error(err);
        alert('加工マスタの読み込みに失敗しました。');
    }
}

function parseProcessingMaster(rows) {
    if (rows.length < 2) return;

    const header = rows[0];
    let colName = header.indexOf('加工方法名');
    if (colName === -1) {
        colName = header.indexOf('加工方法');
    }
    const colId = header.indexOf('加工ID');

    if (colName === -1 || colId === -1) {
        alert('エラー：加工マスタに必要な列（加工方法名 または 加工方法, 加工ID）が見つかりません。');
        return;
    }

    state.processing = {};
    dom.inputProcessing.innerHTML = '<option value="">選択してください</option>';

    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        const name = row[colName];
        const id = row[colId];

        if (name && id) {
            state.processing[name] = id;
            const option = document.createElement('option');
            option.value = name;
            option.textContent = name;
            dom.inputProcessing.appendChild(option);
        }
    }
}

// --- Image Handling ---

function handleImagesLoad(e) {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;

    state.images = files.sort((a, b) => a.name.localeCompare(b.name));
    state.currentIndex = 0;
    state.renamedFiles = {};

    updateStatus(dom.statusImages, true, `画像 ${files.length}枚`);
    dom.fileImages.parentElement.classList.replace('btn-primary', 'btn-success');

    displayCurrentImage();
    updateUIState();
}

function displayCurrentImage() {
    if (state.images.length === 0) {
        dom.previewImage.style.display = 'none';
        dom.placeholder.style.display = 'block';
        dom.imageCounter.textContent = 'No Images';
        dom.currentFilename.textContent = '-';
        return;
    }

    const file = state.images[state.currentIndex];

    // Create Object URL
    const url = URL.createObjectURL(file);
    dom.previewImage.src = url;
    dom.previewImage.style.display = 'block';
    dom.placeholder.style.display = 'none';

    // Update info
    dom.imageCounter.textContent = `Image ${state.currentIndex + 1} / ${state.images.length}`;
    dom.currentFilename.textContent = file.name;

    // Check if already renamed
    if (state.renamedFiles[file.name]) {
        dom.renamePreview.textContent = `✅ Renamed: ${state.renamedFiles[file.name]}`;
        dom.renamePreview.style.color = 'var(--success)';
        // Restore input values if possible (skipped for simplicity, keeping persistence logic)
    } else {
        updateRenamePreview();
    }
}

// --- Logic & UI Updates ---

function handleCategoryChange() {
    const category = dom.inputCategory.value;
    dom.inputMaterial.innerHTML = '<option value="">選択してください</option>';

    if (category && state.materialCategories[category]) {
        state.materialCategories[category].forEach(matName => {
            const option = document.createElement('option');
            option.value = matName;
            option.textContent = matName;
            dom.inputMaterial.appendChild(option);
        });
        dom.inputMaterial.disabled = false;
    } else {
        dom.inputMaterial.disabled = true;
    }
    validateInputs();
}

function updateUIState() {
    const ready = isReady();

    // Enable/Disable Inputs
    const inputs = [
        dom.inputPartName, dom.inputWeight, dom.inputUnit,
        dom.inputCategory, dom.inputProcessing,
        dom.inputPhotoType, dom.inputNotes
    ];
    inputs.forEach(el => el.disabled = !ready);

    // Material Input Logic
    dom.inputMaterial.disabled = !ready || !dom.inputCategory.value;

    // Navigation
    dom.btnPrev.disabled = !ready || state.currentIndex === 0;
    dom.btnNext.disabled = !ready || state.currentIndex === state.images.length - 1;

    // Download
    dom.btnDownload.disabled = Object.keys(state.renamedFiles).length === 0;

    // Focus
    if (ready && document.activeElement === document.body) {
        dom.inputPartName.focus();
    }
}

function validateInputs() {
    const values = getInputValues();
    const isValid = (
        values.partName &&
        values.weight && /^[0-9a-zA-Z.-]+$/.test(values.weight) &&
        values.category &&
        values.material &&
        values.processing
    );

    dom.btnApply.disabled = !isValid;
    return isValid;
}

function getInputValues() {
    return {
        partName: dom.inputPartName.value.trim(),
        weight: dom.inputWeight.value.trim(),
        unit: dom.inputUnit.value,
        category: dom.inputCategory.value,
        material: dom.inputMaterial.value,
        processing: dom.inputProcessing.value,
        photoType: dom.inputPhotoType.value,
        notes: dom.inputNotes.value
    };
}

function sanitizeFilename(str) {
    if (!str) return 'untitled';
    // Remove forbidden chars
    return str.replace(/[<>:"/\\|?*\x00-\x1f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
}

function generateNewFilename() {
    if (!isReady()) return '';

    const v = getInputValues();
    if (!v.partName || !v.material || !v.processing) return '...';

    const partName = sanitizeFilename(v.partName);
    const weight = sanitizeFilename(v.weight);

    // Get IDs
    const matId = state.materialNameToId[v.material] || '???';
    const procId = state.processing[v.processing] || '???';

    // Numbering (Current Index + 1)
    // Note: In local version, it scans folder for max ID.
    // Here, we simply use the array index + 1 for simplicity and sequential processing.
    const number = state.currentIndex + 1;

    // Format: Number_PartName_Weight_Unit_MatID_ProcID_PhotoType_Notes
    return `${number}_${partName}_${weight}_${v.unit}_${matId}_${procId}_${v.photoType}_${v.notes}`;
}

function updateRenamePreview() {
    const currentFile = state.images[state.currentIndex];
    if (!currentFile) return;

    const newName = generateNewFilename();
    if (newName === '...') {
        dom.renamePreview.textContent = 'Enter details to preview filename...';
        dom.renamePreview.style.color = 'var(--text-secondary)';
    } else {
        const ext = currentFile.name.split('.').pop();
        dom.renamePreview.textContent = `➡ ${newName}.${ext}`;
        dom.renamePreview.style.color = 'var(--primary)';
    }
}

function applyAndNext() {
    if (!validateInputs()) return;

    const currentFile = state.images[state.currentIndex];
    const newNameBase = generateNewFilename();
    const ext = currentFile.name.split('.').pop();
    const newFullname = `${newNameBase}.${ext}`;

    // Validated, save to renamed map
    state.renamedFiles[currentFile.name] = newFullname;

    // Next image
    if (state.currentIndex < state.images.length - 1) {
        state.currentIndex++;
        displayCurrentImage();
        updateUIState();

        // Keep Part Name & Weight (Persistence)
        // Focus back to Part Name for workflow efficiency
        dom.inputPartName.focus();
    } else {
        alert('全ての画像の処理が完了しました！\nZIPファイルをダウンロードできます。');
        updateUIState();
    }
}

function navigateImage(delta) {
    const newIndex = state.currentIndex + delta;
    if (newIndex >= 0 && newIndex < state.images.length) {
        state.currentIndex = newIndex;
        displayCurrentImage();
        updateUIState();
    }
}

// --- ZIP Generation ---

async function downloadZip() {
    const zip = new JSZip();
    const folder = zip.folder("renamed_images");
    let count = 0;

    dom.btnDownload.textContent = "⏳ Compressing...";
    dom.btnDownload.disabled = true;

    try {
        for (const file of state.images) {
            const newName = state.renamedFiles[file.name];
            if (newName) {
                folder.file(newName, file);
                count++;
            }
        }

        if (count === 0) {
            alert('リネームされたファイルがありません。');
            return;
        }

        const blob = await zip.generateAsync({ type: "blob" });
        saveAs(blob, "renamed_images.zip");

    } catch (err) {
        console.error(err);
        alert('ZIP作成に失敗しました。');
    } finally {
        dom.btnDownload.textContent = "⬇️ Rename & Download ZIP";
        dom.btnDownload.disabled = false;
    }
}

// --- Utilities ---

function isReady() {
    return (
        Object.keys(state.materials).length > 0 &&
        Object.keys(state.processing).length > 0 &&
        state.images.length > 0
    );
}

function updateStatus(el, success, text) {
    el.textContent = text;
    el.className = `status-badge ${success ? 'success' : 'pending'}`;
}

function resetApp() {
    if (!confirm('全てのデータをリセットしますか？')) return;
    location.reload();
}

// Start
init();
