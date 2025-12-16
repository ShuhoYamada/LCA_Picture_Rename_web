/**
 * Image Renamer Pro - Web版
 * メインアプリケーションロジック
 */

// グローバル状態管理
const AppState = {
    // マスターデータ
    materials: {},
    processingMethods: {},
    implementers: {},
    materialCategories: {},
    materialNameToId: {},
    
    // 画像データ
    imageFiles: [],
    currentIndex: 0,
    processedFiles: new Map(), // originalName -> { newName, blob }
    
    // UI状態
    isReady: false
};

// 初期化
document.addEventListener('DOMContentLoaded', () => {
    initializeApp();
});

function initializeApp() {
    // イベントリスナー設定
    document.getElementById('excelFile').addEventListener('change', handleExcelUpload);
    document.getElementById('imageFiles').addEventListener('change', handleImageUpload);
    document.getElementById('materialCategorySelect').addEventListener('change', handleMaterialCategoryChange);
    document.getElementById('prevButton').addEventListener('click', navigatePrevious);
    document.getElementById('applyButton').addEventListener('click', applyAndNext);
    document.getElementById('downloadButton').addEventListener('click', downloadZip);
    
    // フォーム変更時のプレビュー更新
    const formInputs = ['numberInput', 'implementerSelect', 'partNameInput', 'weightInput', 
                       'unitSelect', 'materialSelect', 'processingSelect', 'photoTypeSelect', 'notesSelect'];
    formInputs.forEach(id => {
        const element = document.getElementById(id);
        if (element) {
            element.addEventListener('input', updateFilenamePreview);
            element.addEventListener('change', updateFilenamePreview);
        }
    });
}

/**
 * Excelファイル読み込み処理
 */
async function handleExcelUpload(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    try {
        updateStatus('excelStatus', 'Excelファイルを読み込み中...', 'info');
        document.getElementById('excelFileName').textContent = file.name;
        
        const data = await readExcelFile(file);
        parseExcelData(data);
        
        updateStatus('excelStatus', `✅ 読み込み完了: 素材${Object.keys(AppState.materials).length}件、加工方法${Object.keys(AppState.processingMethods).length}件、実施者${Object.keys(AppState.implementers).length}件`, 'success');
        
        checkReadyState();
    } catch (error) {
        updateStatus('excelStatus', `❌ エラー: ${error.message}`, 'error');
        console.error(error);
    }
}

/**
 * Excelファイルを読み込む
 */
function readExcelFile(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                resolve(workbook);
            } catch (error) {
                reject(new Error('Excelファイルの解析に失敗しました'));
            }
        };
        
        reader.onerror = () => reject(new Error('ファイルの読み込みに失敗しました'));
        reader.readAsArrayBuffer(file);
    });
}

/**
 * Excelデータを解析
 */
function parseExcelData(workbook) {
    // 素材シート
    const materialSheet = findSheet(workbook, '素材');
    if (materialSheet) {
        parseMaterialSheet(materialSheet);
    }
    
    // 加工方法シート
    const processingSheet = findSheet(workbook, '加工');
    if (processingSheet) {
        parseProcessingSheet(processingSheet);
    }
    
    // 実施者シート
    const implementerSheet = findSheet(workbook, '実施者');
    if (implementerSheet) {
        parseImplementerSheet(implementerSheet);
    }
    
    if (Object.keys(AppState.materials).length === 0 || 
        Object.keys(AppState.processingMethods).length === 0 || 
        Object.keys(AppState.implementers).length === 0) {
        throw new Error('必要なシート（素材、加工、実施者）が見つかりません');
    }
}

/**
 * シート名から該当するシートを検索
 */
function findSheet(workbook, keyword) {
    const sheetName = workbook.SheetNames.find(name => name.includes(keyword));
    return sheetName ? workbook.Sheets[sheetName] : null;
}

/**
 * 素材シートを解析
 */
function parseMaterialSheet(sheet) {
    const data = XLSX.utils.sheet_to_json(sheet);
    AppState.materials = {};
    AppState.materialCategories = {};
    AppState.materialNameToId = {};
    
    data.forEach(row => {
        const name = row['素材名'];
        const id = row['素材ID'];
        const category = row['素材区分'];
        
        if (name && id && category) {
            AppState.materials[name] = id;
            AppState.materialNameToId[name] = id;
            
            if (!AppState.materialCategories[category]) {
                AppState.materialCategories[category] = [];
            }
            AppState.materialCategories[category].push(name);
        }
    });
    
    // 素材区分のドロップダウンを更新
    updateSelect('materialCategorySelect', Object.keys(AppState.materialCategories));
}

/**
 * 加工方法シートを解析
 */
function parseProcessingSheet(sheet) {
    const data = XLSX.utils.sheet_to_json(sheet);
    AppState.processingMethods = {};
    
    data.forEach(row => {
        const name = row['加工方法名'];
        const id = row['加工ID'];
        
        if (name && id) {
            AppState.processingMethods[name] = id;
        }
    });
    
    // 加工方法のドロップダウンを更新
    updateSelect('processingSelect', Object.keys(AppState.processingMethods));
}

/**
 * 実施者シートを解析
 */
function parseImplementerSheet(sheet) {
    const data = XLSX.utils.sheet_to_json(sheet);
    AppState.implementers = {};
    
    data.forEach(row => {
        const name = row['実施者名'] || row['名前'];
        const id = row['実施者ID'] || row['ID'];
        
        if (name && id) {
            AppState.implementers[name] = id;
        }
    });
    
    // 実施者のドロップダウンを更新
    updateSelect('implementerSelect', Object.keys(AppState.implementers));
}

/**
 * セレクトボックスを更新
 */
function updateSelect(selectId, options) {
    const select = document.getElementById(selectId);
    // 最初のオプション（プレースホルダー）以外を削除
    while (select.options.length > 1) {
        select.remove(1);
    }
    
    options.forEach(option => {
        const opt = document.createElement('option');
        opt.value = option;
        opt.textContent = option;
        select.appendChild(opt);
    });
}

/**
 * 素材区分変更時の処理
 */
function handleMaterialCategoryChange(event) {
    const category = event.target.value;
    const materialSelect = document.getElementById('materialSelect');
    
    if (category) {
        const materials = AppState.materialCategories[category] || [];
        updateSelect('materialSelect', materials);
        materialSelect.disabled = false;
    } else {
        materialSelect.disabled = true;
        materialSelect.innerHTML = '<option value="">素材区分を選択してください</option>';
    }
    
    updateFilenamePreview();
}

/**
 * 画像ファイル読み込み処理
 */
async function handleImageUpload(event) {
    const files = Array.from(event.target.files);
    if (files.length === 0) return;
    
    try {
        updateStatus('imageStatus', `${files.length}個の画像ファイルを読み込み中...`, 'info');
        
        // ファイルを自然順序でソート
        files.sort((a, b) => naturalSort(a.name, b.name));
        
        AppState.imageFiles = files;
        AppState.currentIndex = 0;
        AppState.processedFiles.clear();
        
        document.getElementById('imageFileName').textContent = `${files.length}個のファイルを選択`;
        updateStatus('imageStatus', `✅ ${files.length}個の画像ファイルを読み込みました`, 'success');
        
        checkReadyState();
        
        if (AppState.isReady) {
            displayCurrentImage();
            updateNavigationButtons();
            autoSetNumber();
        }
    } catch (error) {
        updateStatus('imageStatus', `❌ エラー: ${error.message}`, 'error');
        console.error(error);
    }
}

/**
 * 自然順序ソート
 */
function naturalSort(a, b) {
    const ax = [], bx = [];
    
    a.replace(/(\d+)|(\D+)/g, (_, $1, $2) => { ax.push([$1 || Infinity, $2 || '']); });
    b.replace(/(\d+)|(\D+)/g, (_, $1, $2) => { bx.push([$1 || Infinity, $2 || '']); });
    
    while (ax.length && bx.length) {
        const an = ax.shift();
        const bn = bx.shift();
        const nn = (an[0] - bn[0]) || an[1].localeCompare(bn[1]);
        if (nn) return nn;
    }
    
    return ax.length - bx.length;
}

/**
 * 現在の画像を表示
 */
function displayCurrentImage() {
    const file = AppState.imageFiles[AppState.currentIndex];
    if (!file) return;
    
    const reader = new FileReader();
    reader.onload = (e) => {
        const preview = document.getElementById('imagePreview');
        preview.innerHTML = `<img src="${e.target.result}" alt="画像プレビュー">`;
        
        const info = document.getElementById('imageInfo');
        info.textContent = `画像 ${AppState.currentIndex + 1} / ${AppState.imageFiles.length}: ${file.name}`;
    };
    
    reader.readAsDataURL(file);
    updateFilenamePreview();
}

/**
 * ナビゲーションボタンの状態更新
 */
function updateNavigationButtons() {
    document.getElementById('prevButton').disabled = AppState.currentIndex === 0;
    updateApplyButtonState();
}

/**
 * 適用ボタンの状態更新
 */
function updateApplyButtonState() {
    const isValid = validateForm();
    document.getElementById('applyButton').disabled = !isValid;
}

/**
 * フォームバリデーション
 */
function validateForm() {
    const partName = document.getElementById('partNameInput').value.trim();
    const weight = document.getElementById('weightInput').value.trim();
    const implementer = document.getElementById('implementerSelect').value;
    const material = document.getElementById('materialSelect').value;
    const processing = document.getElementById('processingSelect').value;
    
    return partName && weight && implementer && material && processing;
}

/**
 * ファイル名プレビュー更新
 */
function updateFilenamePreview() {
    const preview = document.getElementById('filenamePreview');
    
    try {
        const filename = generateFilename();
        if (filename) {
            const file = AppState.imageFiles[AppState.currentIndex];
            const ext = file ? file.name.split('.').pop() : 'jpg';
            preview.textContent = `${filename}.${ext}`;
        } else {
            preview.textContent = 'ファイル名がここに表示されます';
        }
    } catch (error) {
        preview.textContent = 'ファイル名がここに表示されます';
    }
    
    updateApplyButtonState();
}

/**
 * ファイル名を生成
 */
function generateFilename() {
    const number = document.getElementById('numberInput').value || getNextNumber();
    const implementer = document.getElementById('implementerSelect').value;
    const partName = document.getElementById('partNameInput').value.trim();
    const weight = document.getElementById('weightInput').value.trim();
    const unit = document.getElementById('unitSelect').value;
    const material = document.getElementById('materialSelect').value;
    const processing = document.getElementById('processingSelect').value;
    const photoType = document.getElementById('photoTypeSelect').value;
    const notes = document.getElementById('notesSelect').value;
    
    if (!partName || !weight || !implementer || !material || !processing) {
        return '';
    }
    
    const implementerId = AppState.implementers[implementer];
    const materialId = AppState.materialNameToId[material];
    const processingId = AppState.processingMethods[processing];
    
    // ファイル名形式: 番号_部品名_重量_単位_素材ID_加工ID_実施者ID_写真区分_特記事項
    return `${number}_${partName}_${weight}_${unit}_${materialId}_${processingId}_${implementerId}_${photoType}_${notes}`;
}

/**
 * 次の番号を取得（ペアロジック: 1,1,2,2,3,3...）
 */
function getNextNumber() {
    const numberCounts = {};
    
    // 処理済みファイルから番号をカウント
    AppState.processedFiles.forEach((data, originalName) => {
        const match = data.newName.match(/^(\d+)_/);
        if (match) {
            const num = parseInt(match[1]);
            numberCounts[num] = (numberCounts[num] || 0) + 1;
        }
    });
    
    if (Object.keys(numberCounts).length === 0) {
        return 1;
    }
    
    const maxNumber = Math.max(...Object.keys(numberCounts).map(Number));
    
    if (numberCounts[maxNumber] < 2) {
        return maxNumber;
    } else {
        return maxNumber + 1;
    }
}

/**
 * 番号を自動設定
 */
function autoSetNumber() {
    const numberInput = document.getElementById('numberInput');
    numberInput.value = getNextNumber();
    updateFilenamePreview();
}

/**
 * 前の画像へ移動
 */
function navigatePrevious() {
    if (AppState.currentIndex > 0) {
        AppState.currentIndex--;
        displayCurrentImage();
        updateNavigationButtons();
        autoSetNumber();
    }
}

/**
 * 適用して次へ
 */
async function applyAndNext() {
    if (!validateForm()) {
        alert('すべての必須項目を入力してください。');
        return;
    }
    
    try {
        const file = AppState.imageFiles[AppState.currentIndex];
        const newFilename = generateFilename();
        const ext = file.name.split('.').pop();
        const fullNewFilename = `${newFilename}.${ext}`;
        
        // ファイルデータを保存
        AppState.processedFiles.set(file.name, {
            newName: newFilename,
            extension: ext,
            blob: file
        });
        
        // 処理済みリストに追加
        addToProcessedList(file.name, fullNewFilename);
        
        // 次の画像へ
        if (AppState.currentIndex < AppState.imageFiles.length - 1) {
            AppState.currentIndex++;
            displayCurrentImage();
            updateNavigationButtons();
            autoSetNumber();
            
            // 部品名フィールドにフォーカス
            document.getElementById('partNameInput').focus();
        } else {
            // すべて完了
            showCompletionMessage();
        }
        
        // ダウンロードボタンを有効化
        updateDownloadButton();
        
    } catch (error) {
        alert(`エラー: ${error.message}`);
        console.error(error);
    }
}

/**
 * 処理済みリストに追加
 */
function addToProcessedList(originalName, newName) {
    const list = document.getElementById('processedList');
    
    // プレースホルダーを削除
    const placeholder = list.querySelector('.placeholder-text');
    if (placeholder) {
        placeholder.remove();
    }
    
    const item = document.createElement('div');
    item.className = 'processed-item';
    item.textContent = newName;
    list.appendChild(item);
    
    // スクロールを一番下に
    list.scrollTop = list.scrollHeight;
}

/**
 * 完了メッセージを表示
 */
function showCompletionMessage() {
    const preview = document.getElementById('imagePreview');
    preview.innerHTML = `
        <div class="placeholder">
            <span class="placeholder-icon">🎉</span>
            <p><strong>すべての画像の処理が完了しました！</strong><br>
            「リネーム済みファイルをダウンロード」ボタンから<br>
            ZIPファイルをダウンロードできます。</p>
        </div>
    `;
    
    document.getElementById('imageInfo').textContent = `完了: ${AppState.processedFiles.size}個のファイルを処理しました`;
    document.getElementById('applyButton').disabled = true;
}

/**
 * ダウンロードボタンの状態更新
 */
function updateDownloadButton() {
    const button = document.getElementById('downloadButton');
    button.disabled = AppState.processedFiles.size === 0;
    
    if (AppState.processedFiles.size > 0) {
        button.textContent = `💾 リネーム済みファイルをダウンロード (${AppState.processedFiles.size}個のファイル)`;
    }
}

/**
 * ZIPファイルをダウンロード
 */
async function downloadZip() {
    try {
        updateStatus('downloadStatus', 'ZIPファイルを生成中...', 'info');
        
        const zip = new JSZip();
        
        // 処理済みファイルをZIPに追加
        for (const [originalName, data] of AppState.processedFiles) {
            const fullFilename = `${data.newName}.${data.extension}`;
            zip.file(fullFilename, data.blob);
        }
        
        // ZIPを生成
        const content = await zip.generateAsync({ type: 'blob' });
        
        // ダウンロード
        const timestamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        saveAs(content, `renamed_images_${timestamp}.zip`);
        
        updateStatus('downloadStatus', `✅ ${AppState.processedFiles.size}個のファイルをダウンロードしました`, 'success');
        
    } catch (error) {
        updateStatus('downloadStatus', `❌ エラー: ${error.message}`, 'error');
        console.error(error);
    }
}

/**
 * ステータスメッセージを更新
 */
function updateStatus(elementId, message, type = 'info') {
    const element = document.getElementById(elementId);
    element.textContent = message;
    element.className = `status-message ${type}`;
    element.style.display = message ? 'block' : 'none';
}

/**
 * 準備完了状態をチェック
 */
function checkReadyState() {
    const hasExcel = Object.keys(AppState.materials).length > 0 && 
                     Object.keys(AppState.processingMethods).length > 0 &&
                     Object.keys(AppState.implementers).length > 0;
    const hasImages = AppState.imageFiles.length > 0;
    
    AppState.isReady = hasExcel && hasImages;
    
    if (AppState.isReady) {
        // 入力フィールドを有効化
        enableInputFields();
    }
}

/**
 * 入力フィールドを有効化
 */
function enableInputFields() {
    document.getElementById('numberInput').disabled = false;
    document.getElementById('implementerSelect').disabled = false;
    document.getElementById('partNameInput').disabled = false;
    document.getElementById('weightInput').disabled = false;
    document.getElementById('unitSelect').disabled = false;
    document.getElementById('materialCategorySelect').disabled = false;
    document.getElementById('processingSelect').disabled = false;
    document.getElementById('photoTypeSelect').disabled = false;
    document.getElementById('notesSelect').disabled = false;
    
    // 部品名フィールドにフォーカス
    document.getElementById('partNameInput').focus();
}
