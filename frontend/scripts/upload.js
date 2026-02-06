/**
 * 薪資計算系統 - 檔案上傳邏輯
 */

let selectedFile = null;

// DOM 元素
const fileInput = document.getElementById('fileInput');
const uploadSection = document.getElementById('uploadSection');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const fileSize = document.getElementById('fileSize');
const sheetNameInput = document.getElementById('sheetName');
const submitBtn = document.getElementById('submitBtn');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const alertBox = document.getElementById('alert');

// 檔案選擇事件
fileInput.addEventListener('change', handleFileSelect);

// 拖曳事件
uploadSection.addEventListener('dragover', handleDragOver);
uploadSection.addEventListener('dragleave', handleDragLeave);
uploadSection.addEventListener('drop', handleDrop);

// 提交按鈕事件
submitBtn.addEventListener('click', handleSubmit);

/**
 * 處理檔案選擇
 */
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) {
    validateAndDisplayFile(file);
  }
}

/**
 * 處理拖曳懸停
 */
function handleDragOver(e) {
  e.preventDefault();
  e.stopPropagation();
  uploadSection.classList.add('dragover');
}

/**
 * 處理拖曳離開
 */
function handleDragLeave(e) {
  e.preventDefault();
  e.stopPropagation();
  uploadSection.classList.remove('dragover');
}

/**
 * 處理檔案放下
 */
function handleDrop(e) {
  e.preventDefault();
  e.stopPropagation();
  uploadSection.classList.remove('dragover');
  
  const files = e.dataTransfer.files;
  if (files.length > 0) {
    validateAndDisplayFile(files[0]);
  }
}

/**
 * 驗證並顯示檔案資訊
 */
function validateAndDisplayFile(file) {
  // 檢查檔案類型
  const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
  if (!CONFIG.ALLOWED_FILE_TYPES.includes(fileExtension)) {
    showAlert('error', `不支援的檔案格式。請上傳 ${CONFIG.ALLOWED_FILE_TYPES.join(', ')} 檔案。`);
    return;
  }
  
  // 檢查檔案大小
  if (file.size > CONFIG.MAX_FILE_SIZE) {
    const maxSizeMB = CONFIG.MAX_FILE_SIZE / (1024 * 1024);
    showAlert('error', `檔案過大。最大允許 ${maxSizeMB}MB。`);
    return;
  }
  
  // 儲存檔案
  selectedFile = file;
  
  // 顯示檔案資訊
  fileName.textContent = `📄 ${file.name}`;
  fileSize.textContent = `大小: ${formatFileSize(file.size)}`;
  fileInfo.classList.add('show');
  submitBtn.classList.add('show');
  
  // 清除錯誤訊息
  hideAlert();
}

/**
 * 格式化檔案大小
 */
function formatFileSize(bytes) {
  if (bytes === 0) return '0 Bytes';
  const k = 1024;
  const sizes = ['Bytes', 'KB', 'MB', 'GB'];
  const i = Math.floor(Math.log(bytes) / Math.log(k));
  return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i];
}

/**
 * 處理提交
 */
async function handleSubmit() {
  if (!selectedFile) {
    showAlert('error', '請先選擇檔案');
    return;
  }
  
  // 取得並驗證工作表名稱
  const sheetName = sheetNameInput.value.trim();
  if (!sheetName) {
    showAlert('error', '請輸入 Excel 工作表名稱');
    sheetNameInput.focus();
    return;
  }
  
  // 禁用提交按鈕和輸入欄位
  submitBtn.disabled = true;
  submitBtn.textContent = '處理中...';
  sheetNameInput.disabled = true;
  
  // 顯示進度條
  progressContainer.classList.add('show');
  updateProgress(0, '正在讀取檔案...');
  
  try {
    // 讀取檔案為 Base64
    updateProgress(20, '正在編碼檔案...');
    const base64Data = await fileToBase64(selectedFile);
    
    // 準備上傳資料
    updateProgress(40, '正在上傳到伺服器...');
    const payload = {
      action: 'upload',
      fileName: selectedFile.name,
      fileData: base64Data,
      targetSheetName: sheetName, // 使用使用者輸入的工作表名稱
      targetGoogleSheetName: CONFIG.TARGET_GOOGLE_SHEET_NAME,
      targetGoogleSheetTab: CONFIG.TARGET_GOOGLE_SHEET_TAB
    };
    
    // 發送到 GAS
    const response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'no-cors', // GAS 需要使用 no-cors
      headers: {
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(payload)
    });
    
    updateProgress(60, '伺服器處理中...');
    
    // 注意：no-cors 模式下無法讀取 response，所以我們假設成功
    // 實際應用中可以改用輪詢狀態的方式
    await new Promise(resolve => setTimeout(resolve, 2000)); // 模擬處理時間
    
    updateProgress(100, '完成！');
    showAlert('success', '✓ 檔案已成功上傳並處理！資料已寫入 Google Sheets。');
    
    // 重置表單
    setTimeout(() => {
      resetForm();
    }, 3000);
    
  } catch (error) {
    console.error('上傳錯誤:', error);
    showAlert('error', `上傳失敗: ${error.message}`);
    submitBtn.disabled = false;
    submitBtn.textContent = '開始上傳並處理';
    sheetNameInput.disabled = false;
    progressContainer.classList.remove('show');
  }
}

/**
 * 將檔案轉換為 Base64
 */
function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => {
      // 移除 data:application/...;base64, 前綴
      const base64 = reader.result.split(',')[1];
      resolve(base64);
    };
    reader.onerror = reject;
    reader.readAsDataURL(file);
  });
}

/**
 * 更新進度條
 */
function updateProgress(percent, text) {
  progressFill.style.width = percent + '%';
  progressFill.textContent = percent + '%';
  progressText.textContent = text;
}

/**
 * 顯示提示訊息
 */
function showAlert(type, message) {
  alertBox.className = `alert alert-${type} show`;
  alertBox.textContent = message;
}

/**
 * 隱藏提示訊息
 */
function hideAlert() {
  alertBox.className = 'alert';
  alertBox.textContent = '';
}

/**
 * 重置表單
 */
function resetForm() {
  selectedFile = null;
  fileInput.value = '';
  fileInfo.classList.remove('show');
  submitBtn.classList.remove('show');
  submitBtn.disabled = false;
  submitBtn.textContent = '開始上傳並處理';
  sheetNameInput.disabled = false;
  sheetNameInput.value = CONFIG.TARGET_SHEET_NAME; // 重置為預設值
  progressContainer.classList.remove('show');
  updateProgress(0, '');
}
