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
const sheetNameHint = document.getElementById('sheetNameHint');
const submitBtn = document.getElementById('submitBtn');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const alertBox = document.getElementById('alert');
const resultSection = document.getElementById('resultSection');
const resultSummary = document.getElementById('resultSummary');
const resultList = document.getElementById('resultList');
const yearSelect = document.getElementById('yearSelect');
const monthSelect = document.getElementById('monthSelect');
const yearMonthInput = document.getElementById('yearMonthInput');
const datePicker = document.getElementById('datePicker');
const loadScheduleBtn = document.getElementById('loadScheduleBtn');
const scheduleResultSection = document.getElementById('scheduleResultSection');
const scheduleSummary = document.getElementById('scheduleSummary');
const scheduleList = document.getElementById('scheduleList');
const personCheckboxGroup = document.getElementById('personCheckboxGroup');
const selectAllPersonsBtn = document.getElementById('selectAllPersonsBtn');
const clearAllPersonsBtn = document.getElementById('clearAllPersonsBtn');
const branchSelect = document.getElementById('branchSelect');
const branchGroup = document.getElementById('branchGroup');

// 初始化年月選擇器
function initScheduleSelectors() {
  const now = new Date();
  const currentYear = now.getFullYear();
  for (let y = currentYear - 2; y <= currentYear + 2; y++) {
    const opt = document.createElement('option');
    opt.value = String(y);
    opt.textContent = String(y);
    if (y === currentYear) opt.selected = true;
    yearSelect.appendChild(opt);
  }
  for (let m = 1; m <= 12; m++) {
    const opt = document.createElement('option');
    opt.value = String(m).padStart(2, '0');
    opt.textContent = String(m).padStart(2, '0') + ' 月';
    if (m === now.getMonth() + 1) opt.selected = true;
    monthSelect.appendChild(opt);
  }
}

// 檔案選擇事件
fileInput.addEventListener('change', handleFileSelect);

// 拖曳事件
uploadSection.addEventListener('dragover', handleDragOver);
uploadSection.addEventListener('dragleave', handleDragLeave);
uploadSection.addEventListener('drop', handleDrop);

// 提交按鈕事件
submitBtn.addEventListener('click', handleSubmit);

// 上傳類型切換（更新工作表名稱說明、分店區塊顯示）
document.querySelectorAll('input[name="uploadType"]').forEach(function(radio) {
  radio.addEventListener('change', function() {
    const isSchedule = this.value === 'schedule';
    const isAttendance = this.value === 'attendance';
    if (sheetNameHint) {
      sheetNameHint.textContent = isSchedule
        ? '請輸入要處理的 Excel 工作表名稱（例如：11501、11502）'
        : (isAttendance ? '打卡上傳 CSV 不需輸入工作表名稱' : '請輸入要處理的 Excel 工作表名稱（例如：打卡紀錄、Sheet1）');
    }
    if (sheetNameInput && !sheetNameInput.value) {
      sheetNameInput.placeholder = isSchedule ? '例如：11501' : (isAttendance ? '不需輸入' : '例如：打卡紀錄');
    }
    if (branchGroup) branchGroup.style.display = (isSchedule || isAttendance) ? 'block' : 'none';
    if (selectedFile) {
      const ext = '.' + selectedFile.name.split('.').pop().toLowerCase();
      const ok = isAttendance ? ext === '.csv' : CONFIG.ALLOWED_FILE_TYPES.includes(ext);
      if (!ok) {
        selectedFile = null;
        fileInput.value = '';
        fileInfo.classList.remove('show');
        submitBtn.classList.remove('show');
        showAlert('error', isAttendance ? '打卡請上傳 .csv 檔案，請重新選擇' : '班表請上傳 .xlsx 或 .xls 檔案，請重新選擇');
      }
    }
  });
});

// 載入按鈕（班表／打卡共用）
if (loadScheduleBtn) loadScheduleBtn.addEventListener('click', handleLoadQuery);

// 日期篩選模式切換
document.querySelectorAll('input[name="dateFilterMode"]').forEach(radio => {
  if (radio) radio.addEventListener('change', toggleDateFilterMode);
});

// 查詢類型切換（班表／打卡）
document.querySelectorAll('input[name="queryType"]').forEach(radio => {
  if (radio) radio.addEventListener('change', function() {
    const titleEl = document.getElementById('querySectionTitle');
    const resultTitleEl = document.getElementById('queryResultTitle');
    if (this.value === 'attendance') {
      if (titleEl) titleEl.textContent = '📅 查詢打卡';
      if (resultTitleEl) resultTitleEl.textContent = '打卡資料';
    } else {
      if (titleEl) titleEl.textContent = '📅 查詢班表';
      if (resultTitleEl) resultTitleEl.textContent = '班表資料';
    }
  });
});
if (datePicker) datePicker.addEventListener('change', function() {});

// 人員篩選按鈕
if (selectAllPersonsBtn) selectAllPersonsBtn.addEventListener('click', selectAllPersons);
if (clearAllPersonsBtn) clearAllPersonsBtn.addEventListener('click', clearAllPersons);

// 比對區塊日期模式切換
document.querySelectorAll('input[name="compareDateMode"]').forEach(function(radio) {
  if (radio) radio.addEventListener('change', toggleCompareDateMode);
});

// 載入比對按鈕
const loadCompareBtn = document.getElementById('loadCompareBtn');
if (loadCompareBtn) loadCompareBtn.addEventListener('click', handleLoadCompare);

// 頁面載入時初始化
document.addEventListener('DOMContentLoaded', function() {
  if (yearSelect && monthSelect) initScheduleSelectors();
  toggleDateFilterMode();
  toggleCompareDateMode();
  loadBranches();
  // 初始顯示分店區塊（班表為預設）
  if (branchGroup) {
    const mode = document.querySelector('input[name="uploadType"]:checked');
    branchGroup.style.display = mode && mode.value === 'schedule' ? 'block' : 'none';
  }
});

/**
 * 載入分店清單（從 GAS getBranches API），供上傳、查詢、比對區塊使用
 */
async function loadBranches() {
  const branchEl = document.getElementById('branchSelect');
  const queryBranchEl = document.getElementById('queryBranchSelect');
  const compareBranchEl = document.getElementById('compareBranchSelect');
  if (!branchEl && !queryBranchEl && !compareBranchEl) return;
  try {
    const response = await fetch(CONFIG.GAS_URL + '?action=getBranches', { method: 'GET', mode: 'cors' });
    const result = await response.json();
    const options = result && result.success && Array.isArray(result.names) && result.names.length > 0
      ? result.names
      : [];
    if (branchEl) {
      branchEl.innerHTML = '<option value="">請選擇分店</option>';
      options.forEach(function(name) {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        branchEl.appendChild(opt);
      });
    }
    if (queryBranchEl) {
      queryBranchEl.innerHTML = '<option value="">全部</option>';
      options.forEach(function(name) {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        queryBranchEl.appendChild(opt);
      });
    }
    if (compareBranchEl) {
      compareBranchEl.innerHTML = '<option value="">請選擇分店</option>';
      options.forEach(function(name) {
        const opt = document.createElement('option');
        opt.value = name;
        opt.textContent = name;
        compareBranchEl.appendChild(opt);
      });
    }
  } catch (error) {
    console.error('載入分店清單失敗:', error);
    if (branchEl) branchEl.innerHTML = '<option value="">載入失敗，請重整頁面</option>';
    if (queryBranchEl) queryBranchEl.innerHTML = '<option value="">載入失敗</option>';
    if (compareBranchEl) compareBranchEl.innerHTML = '<option value="">載入失敗</option>';
  }
}

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
  // 檢查檔案類型（班表：xlsx/xls；打卡：csv）
  const fileExtension = '.' + file.name.split('.').pop().toLowerCase();
  const uploadType = document.querySelector('input[name="uploadType"]:checked');
  const isAttendance = uploadType && uploadType.value === 'attendance';
  const allowedTypes = isAttendance ? ['.csv'] : CONFIG.ALLOWED_FILE_TYPES;
  if (!allowedTypes.includes(fileExtension)) {
    showAlert('error', isAttendance ? '打卡上傳請使用 .csv 檔案' : `班表上傳請使用 ${CONFIG.ALLOWED_FILE_TYPES.join('、')} 檔案`);
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

  // 只要重新選擇檔案，就解除處理中鎖定
  submitBtn.disabled = false;
  submitBtn.textContent = '開始上傳並處理';
  sheetNameInput.disabled = false;
  progressContainer.classList.remove('show');
  clearResults();
  
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
 * 測試 GAS 連線
 */
async function testGASConnection() {
  try {
    const response = await fetch(CONFIG.GAS_URL + '?action=test', {
      method: 'GET',
      mode: 'cors',
      signal: AbortSignal.timeout(15000) // 15 秒逾時，避免 GAS 冷啟動被誤判
    });
    const result = await response.json();
    return response.ok && result && result.success;
  } catch (error) {
    return false;
  }
}

/**
 * 處理提交
 */
async function handleSubmit() {
  if (!selectedFile) {
    showAlert('error', '請先選擇檔案');
    return;
  }
  
  // 班表／打卡上傳時驗證分店必選
  const uploadType = document.querySelector('input[name="uploadType"]:checked');
  const branchName = branchSelect ? branchSelect.value.trim() : '';
  if (uploadType && (uploadType.value === 'schedule' || uploadType.value === 'attendance')) {
    if (!branchName) {
      showAlert('error', '請選擇分店');
      return;
    }
  }

  // 班表需工作表名稱；打卡不需（CSV 為整檔）
  const sheetName = sheetNameInput.value.trim();
  if (uploadType && uploadType.value === 'schedule' && !sheetName) {
    showAlert('error', '請輸入 Excel 工作表名稱');
    return;
  }
  
  // 禁用提交按鈕和輸入欄位
  submitBtn.disabled = true;
  submitBtn.textContent = '處理中...';
  sheetNameInput.disabled = true;
  
  // 顯示進度條
  progressContainer.classList.add('show');
  updateProgress(0, '正在連線到伺服器...');
  clearResults();
  
  // 測試連線（可選，但能提早發現明顯的問題）
  const isConnected = await testGASConnection();
  if (!isConnected) {
    showAlert('warning', '⚠️ 警告：無法連線到伺服器，但仍會嘗試上傳。請確認 GAS URL 是否正確設定。');
    // 給使用者 3 秒時間看到警告
    await new Promise(resolve => setTimeout(resolve, 3000));
  }
  
  updateProgress(10, '正在讀取檔案...');
  
  try {
    // 讀取檔案為 Base64
    updateProgress(20, '正在編碼檔案...');
    const base64Data = await fileToBase64(selectedFile);
    
    // 準備上傳資料
    updateProgress(40, '正在上傳到伺服器...');
    const uploadType = document.querySelector('input[name="uploadType"]:checked');
    const branchName = branchSelect ? branchSelect.value.trim() : '';
    const payload = {
      action: 'upload',
      uploadType: uploadType ? uploadType.value : 'schedule',
      fileName: selectedFile.name,
      fileData: base64Data,
      targetSheetName: sheetName || '',
      targetGoogleSheetName: CONFIG.TARGET_GOOGLE_SHEET_NAME,
      targetGoogleSheetTab: uploadType && uploadType.value === 'attendance' ? '打卡' : CONFIG.TARGET_GOOGLE_SHEET_TAB,
      branchName: (uploadType && (uploadType.value === 'schedule' || uploadType.value === 'attendance')) ? branchName : ''
    };
    
    // 發送到 GAS（改用可讀取回應）
    const response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      body: JSON.stringify(payload)
    });
    
    updateProgress(60, '伺服器處理中...');
    const result = await response.json();
    
    if (!response.ok || !result.success) {
      throw new Error(result.error || '上傳處理失敗');
    }
    
    updateProgress(100, '上傳完成');
    
    // 重新開放再次上傳與輸入工作表名稱
    submitBtn.disabled = false;
    submitBtn.textContent = '開始上傳並處理';
    sheetNameInput.disabled = false;
    
    renderResults(result);
    showAlert('success', '✅ 解析完成，結果如下', { scrollTo: resultSection });
    
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
 * @param {string} type - success | error | warning
 * @param {string} message - 訊息內容
 * @param {Object} options - { scrollTo: Element } 成功時要捲動到的區塊
 */
function showAlert(type, message, options) {
  alertBox.className = `alert alert-${type} show`;
  alertBox.textContent = message;
  if (type === 'error' || type === 'warning') {
    alertBox.scrollIntoView({ behavior: 'auto', block: 'start' });
  } else if (type === 'success' && options && options.scrollTo) {
    options.scrollTo.scrollIntoView({ behavior: 'auto', block: 'start' });
  }
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
  sheetNameInput.value = ''; // 不再帶預設值，改由使用者每次輸入
  progressContainer.classList.remove('show');
  updateProgress(0, '');
  clearResults();
}

/**
 * 顯示解析結果（摘要 + 卡片列表）
 */
function renderResults(result) {
  const details = result.details || {};
  const records = Array.isArray(result.records) ? result.records : [];
  const shiftCodes = Array.isArray(details.shiftCodes) ? details.shiftCodes : [];

  const isAttendanceResult = result.columns && result.columns[0] === '分店' && result.columns[1] === '員工編號';
  const summaryItems = isAttendanceResult
    ? [
        { label: '新增筆數', value: details.rowCount ?? '—' },
        { label: '略過重複', value: details.skippedCount ?? 0 },
        { label: '原始筆數', value: details.parsedRowCount ?? records.length ?? '—' },
        { label: '處理時間', value: details.processTime ? `${details.processTime}s` : '—' },
        { label: '目標工作表', value: details.targetSheet || '—' }
      ]
    : [
        { label: '新增筆數', value: details.rowCount ?? '—' },
        { label: '略過重複', value: details.skippedCount ?? 0 },
        { label: '原始筆數', value: details.parsedRowCount ?? records.length ?? '—' },
        { label: '員工數', value: details.totalEmployees || 0 },
        { label: '班別代碼', value: shiftCodes.length ? shiftCodes.join(', ') : '—' },
        { label: '處理時間', value: details.processTime ? `${details.processTime}s` : '—' },
        { label: '來源工作表', value: details.sourceSheet || '—' },
        { label: '目標工作表', value: details.targetSheet || '—' }
      ];

  resultSummary.innerHTML = summaryItems.map(item => `
    <div class="summary-item">
      <div class="summary-label">${item.label}</div>
      <div class="summary-value">${item.value}</div>
    </div>
  `).join('');

  // 相容後端格式：records 為 { row, isDuplicate } 或舊版純陣列
  const recordList = records.map(r => {
    if (r && typeof r === 'object' && Array.isArray(r.row)) {
      return { row: r.row, isDuplicate: !!r.isDuplicate };
    }
    const row = Array.isArray(r) ? r : [];
    return { row, isDuplicate: false };
  });

  const isAttendance = result.columns && result.columns[0] === '分店' && result.columns[1] === '員工編號';
  resultList.innerHTML = recordList.map(({ row, isDuplicate }) => {
    const duplicateBadge = isDuplicate ? '<span class="result-card-duplicate" title="此筆為重複，已略過寫入">重複</span>' : '';
    if (isAttendance) {
      const branch = row[0];
      const empNo = row[1];
      const empAccount = row[2];
      const name = row[3];
      const date = row[4];
      const start = row[5];
      const end = row[6];
      const hours = row[7];
      const status = row[8];
      return `
      <div class="result-card ${isDuplicate ? 'result-card--duplicate' : ''}">
        ${duplicateBadge}
        <div class="result-row"><span class="result-label">員工編號</span><span class="result-value">${empNo || '—'}</span></div>
        <div class="result-row"><span class="result-label">姓名</span><span class="result-value">${name || '—'}</span></div>
        <div class="result-row"><span class="result-label">員工帳號</span><span class="result-value">${empAccount || '—'}</span></div>
        <div class="result-row"><span class="result-label">打卡日期</span><span class="result-value">${date || '—'}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">分店</span><span class="result-value">${branch || '—'}</span></div>
        <div class="result-row"><span class="result-label">工作時數</span><span class="result-value">${hours || '—'}</span></div>
        <div class="result-row"><span class="result-label">狀態</span><span class="result-value">${status || '—'}</span></div>
      </div>
    `;
    }
    const name = row[0];
    const date = row[1];
    const start = row[2];
    const end = row[3];
    const hours = row[4];
    const shift = row[5];
    const branch = row[6];
    return `
      <div class="result-card ${isDuplicate ? 'result-card--duplicate' : ''}">
        ${duplicateBadge}
        <div class="result-row"><span class="result-label">姓名</span><span class="result-value">${name || '—'}</span></div>
        <div class="result-row"><span class="result-label">日期</span><span class="result-value">${date || '—'}</span></div>
        <div class="result-row"><span class="result-label">班別</span><span class="result-value">${shift || '—'}</span></div>
        <div class="result-row"><span class="result-label">分店</span><span class="result-value">${branch || '—'}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">時數</span><span class="result-value">${hours || '—'}</span></div>
      </div>
    `;
  }).join('');

  resultSection.classList.add('show');
}

/**
 * 清除結果區塊
 */
function clearResults() {
  resultSummary.innerHTML = '';
  resultList.innerHTML = '';
  resultSection.classList.remove('show');
}

function toggleDateFilterMode() {
  const mode = document.querySelector('input[name="dateFilterMode"]:checked');
  const isMonth = mode && mode.value === 'month';
  const dateMonthGroup = document.getElementById('dateMonthGroup');
  const dateDayGroup = document.getElementById('dateDayGroup');
  if (yearSelect) yearSelect.disabled = !isMonth;
  if (monthSelect) monthSelect.disabled = !isMonth;
  if (yearMonthInput) yearMonthInput.disabled = !isMonth;
  if (datePicker) datePicker.disabled = isMonth;
  if (dateMonthGroup) dateMonthGroup.classList.toggle('hidden', !isMonth);
  if (dateDayGroup) dateDayGroup.classList.toggle('hidden', isMonth);
}

function toggleCompareDateMode() {
  const mode = document.querySelector('input[name="compareDateMode"]:checked');
  const isMonth = mode && mode.value === 'month';
  const compareMonthGroup = document.getElementById('compareMonthGroup');
  const compareRangeGroup = document.getElementById('compareRangeGroup');
  const compareYearMonthInput = document.getElementById('compareYearMonthInput');
  const compareStartDate = document.getElementById('compareStartDate');
  const compareEndDate = document.getElementById('compareEndDate');
  if (compareMonthGroup) compareMonthGroup.classList.toggle('hidden', !isMonth);
  if (compareRangeGroup) compareRangeGroup.classList.toggle('hidden', isMonth);
  if (compareYearMonthInput) compareYearMonthInput.disabled = !isMonth;
  if (compareStartDate) compareStartDate.disabled = isMonth;
  if (compareEndDate) compareEndDate.disabled = isMonth;
}

function selectAllPersons() {
  if (!personCheckboxGroup) return;
  personCheckboxGroup.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = true; });
}

function clearAllPersons() {
  if (!personCheckboxGroup) return;
  personCheckboxGroup.querySelectorAll('input[type="checkbox"]').forEach(cb => { cb.checked = false; });
}

function getSelectedPersonNames() {
  if (!personCheckboxGroup) return [];
  const names = [];
  personCheckboxGroup.querySelectorAll('input[type="checkbox"]:checked').forEach(cb => {
    if (cb.value) names.push(cb.value);
  });
  return names;
}

/**
 * 載入查詢（班表或打卡，依年月/日期、分店、人員篩選）
 */
async function handleLoadQuery() {
  const queryType = document.querySelector('input[name="queryType"]:checked');
  const isAttendance = queryType && queryType.value === 'attendance';
  if (isAttendance) {
    return handleLoadAttendance();
  }
  return handleLoadSchedule();
}

/**
 * 載入班表（依年月/日期、分店、人員篩選，AND 關係）
 */
async function handleLoadSchedule() {
  const mode = document.querySelector('input[name="dateFilterMode"]:checked');
  let yearMonth = '';
  let dateParam = '';

  if (mode && mode.value === 'day' && datePicker && datePicker.value) {
    dateParam = datePicker.value;
  } else {
    if (yearMonthInput && yearMonthInput.value.trim().match(/^\d{6}$/)) {
      yearMonth = yearMonthInput.value.trim();
    } else if (yearSelect && monthSelect) {
      yearMonth = yearSelect.value + monthSelect.value;
    }
  }

  if (!yearMonth && !dateParam) {
    showAlert('error', '請選擇整月（年月）或單日');
    return;
  }

  loadScheduleBtn.disabled = true;
  loadScheduleBtn.textContent = '載入中...';
  hideAlert();
  scheduleResultSection.classList.remove('show');

  const names = getSelectedPersonNames();
  const queryBranchEl = document.getElementById('queryBranchSelect');
  const branchVal = queryBranchEl && queryBranchEl.value ? queryBranchEl.value.trim() : '';
  let url = `${CONFIG.GAS_URL}?action=loadSchedule`;
  if (yearMonth) url += `&yearMonth=${encodeURIComponent(yearMonth)}`;
  if (dateParam) url += `&date=${encodeURIComponent(dateParam)}`;
  if (branchVal) url += `&branch=${encodeURIComponent(branchVal)}`;
  if (names.length > 0) url += `&names=${encodeURIComponent(names.join(','))}`;

  try {
    const response = await fetch(url, { method: 'GET', mode: 'cors' });
    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.error || '載入失敗');
    }

    renderScheduleResults(result);
    scheduleResultSection.scrollIntoView({ behavior: 'auto', block: 'start' });
  } catch (error) {
    showAlert('error', '載入班表失敗：' + error.message);
  } finally {
    loadScheduleBtn.disabled = false;
    loadScheduleBtn.textContent = '載入';
  }
}

/**
 * 載入打卡（依年月/日期、分店、人員篩選，AND 關係）
 */
async function handleLoadAttendance() {
  const mode = document.querySelector('input[name="dateFilterMode"]:checked');
  let yearMonth = '';
  let dateParam = '';

  if (mode && mode.value === 'day' && datePicker && datePicker.value) {
    dateParam = datePicker.value;
  } else {
    if (yearMonthInput && yearMonthInput.value.trim().match(/^\d{6}$/)) {
      yearMonth = yearMonthInput.value.trim();
    } else if (yearSelect && monthSelect) {
      yearMonth = yearSelect.value + monthSelect.value;
    }
  }

  if (!yearMonth && !dateParam) {
    showAlert('error', '請選擇整月（年月）或單日');
    return;
  }

  loadScheduleBtn.disabled = true;
  loadScheduleBtn.textContent = '載入中...';
  hideAlert();
  scheduleResultSection.classList.remove('show');

  const names = getSelectedPersonNames();
  const queryBranchEl = document.getElementById('queryBranchSelect');
  const branchVal = queryBranchEl && queryBranchEl.value ? queryBranchEl.value.trim() : '';
  let url = `${CONFIG.GAS_URL}?action=loadAttendance`;
  if (yearMonth) url += `&yearMonth=${encodeURIComponent(yearMonth)}`;
  if (dateParam) url += `&date=${encodeURIComponent(dateParam)}`;
  if (branchVal) url += `&branch=${encodeURIComponent(branchVal)}`;
  if (names.length > 0) url += `&names=${encodeURIComponent(names.join(','))}`;

  try {
    const response = await fetch(url, { method: 'GET', mode: 'cors' });
    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.error || '載入失敗');
    }

    renderAttendanceResults(result);
    scheduleResultSection.scrollIntoView({ behavior: 'auto', block: 'start' });
  } catch (error) {
    showAlert('error', '載入打卡失敗：' + error.message);
  } finally {
    loadScheduleBtn.disabled = false;
    loadScheduleBtn.textContent = '載入';
  }
}

/**
 * 顯示班表查詢結果
 */
function renderScheduleResults(result) {
  const details = result.details || {};
  const records = Array.isArray(result.records) ? result.records : [];
  const names = Array.isArray(details.names) ? details.names : [];

  const branchLabel = details.branch ? details.branch : '全部';
  scheduleSummary.innerHTML = `
    <div class="summary-item">
      <div class="summary-label">日期範圍</div>
      <div class="summary-value">${details.date ? details.date.replace(/-/g, '/') : (details.yearMonth ? details.yearMonth.substring(0,4) + '/' + details.yearMonth.substring(4,6) : '—')}</div>
    </div>
    <div class="summary-item">
      <div class="summary-label">分店</div>
      <div class="summary-value">${branchLabel}</div>
    </div>
    <div class="summary-item">
      <div class="summary-label">筆數</div>
      <div class="summary-value">${details.rowCount ?? records.length ?? 0}</div>
    </div>
  `;

  if (personCheckboxGroup) {
    personCheckboxGroup.innerHTML = '';
    if (names.length > 0) {
      names.forEach(n => {
        const label = document.createElement('label');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = n;
        label.appendChild(cb);
        label.appendChild(document.createTextNode(n));
        personCheckboxGroup.appendChild(label);
      });
      if (selectAllPersonsBtn) selectAllPersonsBtn.disabled = false;
      if (clearAllPersonsBtn) clearAllPersonsBtn.disabled = false;
    } else {
      personCheckboxGroup.innerHTML = '<span class="person-placeholder">此範圍無人員資料</span>';
      if (selectAllPersonsBtn) selectAllPersonsBtn.disabled = true;
      if (clearAllPersonsBtn) clearAllPersonsBtn.disabled = true;
    }
  }

  scheduleList.innerHTML = records.map(row => {
    const name = row[0];
    const date = row[1];
    const start = row[2];
    const end = row[3];
    const hours = row[4];
    const shift = row[5];
    const branch = row[6];
    return `
      <div class="result-card">
        <div class="result-row"><span class="result-label">姓名</span><span class="result-value">${name || '—'}</span></div>
        <div class="result-row"><span class="result-label">日期</span><span class="result-value">${date || '—'}</span></div>
        <div class="result-row"><span class="result-label">班別</span><span class="result-value">${shift || '—'}</span></div>
        <div class="result-row"><span class="result-label">分店</span><span class="result-value">${branch || '—'}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">時數</span><span class="result-value">${hours || '—'}</span></div>
      </div>
    `;
  }).join('');

  scheduleResultSection.classList.add('show');
}

/**
 * 顯示打卡查詢結果
 */
function renderAttendanceResults(result) {
  const details = result.details || {};
  const records = Array.isArray(result.records) ? result.records : [];
  const names = Array.isArray(details.names) ? details.names : [];

  const branchLabel = details.branch ? details.branch : '全部';
  scheduleSummary.innerHTML = `
    <div class="summary-item">
      <div class="summary-label">日期範圍</div>
      <div class="summary-value">${details.date ? details.date.replace(/-/g, '/') : (details.yearMonth ? details.yearMonth.substring(0,4) + '/' + details.yearMonth.substring(4,6) : '—')}</div>
    </div>
    <div class="summary-item">
      <div class="summary-label">分店</div>
      <div class="summary-value">${branchLabel}</div>
    </div>
    <div class="summary-item">
      <div class="summary-label">筆數</div>
      <div class="summary-value">${details.rowCount ?? records.length ?? 0}</div>
    </div>
  `;

  if (personCheckboxGroup) {
    personCheckboxGroup.innerHTML = '';
    if (names.length > 0) {
      names.forEach(n => {
        const label = document.createElement('label');
        const cb = document.createElement('input');
        cb.type = 'checkbox';
        cb.value = n;
        label.appendChild(cb);
        label.appendChild(document.createTextNode(n));
        personCheckboxGroup.appendChild(label);
      });
      if (selectAllPersonsBtn) selectAllPersonsBtn.disabled = false;
      if (clearAllPersonsBtn) clearAllPersonsBtn.disabled = false;
    } else {
      personCheckboxGroup.innerHTML = '<span class="person-placeholder">此範圍無人員資料</span>';
      if (selectAllPersonsBtn) selectAllPersonsBtn.disabled = true;
      if (clearAllPersonsBtn) clearAllPersonsBtn.disabled = true;
    }
  }

  scheduleList.innerHTML = records.map(row => {
    const branch = row[0];
    const empNo = row[1];
    const empAccount = row[2];
    const name = row[3];
    const date = row[4];
    const start = row[5];
    const end = row[6];
    const hours = row[7];
    const status = row[8];
    return `
      <div class="result-card">
        <div class="result-row"><span class="result-label">分店</span><span class="result-value">${branch || '—'}</span></div>
        <div class="result-row"><span class="result-label">員工編號</span><span class="result-value">${empNo || '—'}</span></div>
        <div class="result-row"><span class="result-label">員工帳號</span><span class="result-value">${empAccount || '—'}</span></div>
        <div class="result-row"><span class="result-label">姓名</span><span class="result-value">${name || '—'}</span></div>
        <div class="result-row"><span class="result-label">打卡日期</span><span class="result-value">${date || '—'}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">工作時數</span><span class="result-value">${hours || '—'}</span></div>
        <div class="result-row"><span class="result-label">狀態</span><span class="result-value">${status || '—'}</span></div>
      </div>
    `;
  }).join('');

  scheduleResultSection.classList.add('show');
}

/**
 * 載入班表與打卡比對
 */
async function handleLoadCompare() {
  const mode = document.querySelector('input[name="compareDateMode"]:checked');
  const isMonth = mode && mode.value === 'month';
  let yearMonth = '';
  let startDate = '';
  let endDate = '';

  if (isMonth) {
    const compareYearMonthInput = document.getElementById('compareYearMonthInput');
    yearMonth = compareYearMonthInput && compareYearMonthInput.value.trim().match(/^\d{6}$/)
      ? compareYearMonthInput.value.trim()
      : '';
  } else {
    const compareStartDate = document.getElementById('compareStartDate');
    const compareEndDate = document.getElementById('compareEndDate');
    startDate = compareStartDate && compareStartDate.value ? compareStartDate.value.trim() : '';
    endDate = compareEndDate && compareEndDate.value ? compareEndDate.value.trim() : startDate;
  }

  if (!yearMonth && (!startDate || startDate.length !== 10)) {
    showAlert('error', '請選擇月份（例如 202601）或日期區間');
    return;
  }

  const compareBranchSelect = document.getElementById('compareBranchSelect');
  const branchVal = compareBranchSelect && compareBranchSelect.value ? compareBranchSelect.value.trim() : '';
  if (!branchVal) {
    showAlert('error', '請選擇分店');
    return;
  }

  const comparePersonCheckboxGroup = document.getElementById('comparePersonCheckboxGroup');
  const names = [];
  if (comparePersonCheckboxGroup) {
    comparePersonCheckboxGroup.querySelectorAll('input[type="checkbox"]:checked').forEach(function(cb) {
      if (cb.value) names.push(cb.value);
    });
  }

  const loadCompareBtn = document.getElementById('loadCompareBtn');
  const compareResultSection = document.getElementById('compareResultSection');
  if (loadCompareBtn) loadCompareBtn.disabled = true;
  loadCompareBtn.textContent = '載入中...';
  hideAlert();
  if (compareResultSection) compareResultSection.classList.remove('show');

  let url = CONFIG.GAS_URL + '?action=loadCompare&branch=' + encodeURIComponent(branchVal);
  if (yearMonth) url += '&yearMonth=' + encodeURIComponent(yearMonth);
  if (startDate) url += '&startDate=' + encodeURIComponent(startDate);
  if (endDate) url += '&endDate=' + encodeURIComponent(endDate);
  if (names.length > 0) url += '&names=' + encodeURIComponent(names.join(','));

  try {
    const response = await fetch(url, { method: 'GET', mode: 'cors' });
    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.error || '載入失敗');
    }

    renderCompareResults(result.items || []);
    populateComparePersonCheckboxes(result.items || []);
    if (compareResultSection) {
      compareResultSection.classList.add('show');
      compareResultSection.scrollIntoView({ behavior: 'auto', block: 'start' });
    }
  } catch (error) {
    showAlert('error', '載入比對失敗：' + error.message);
  } finally {
    if (loadCompareBtn) {
      loadCompareBtn.disabled = false;
      loadCompareBtn.textContent = '載入比對';
    }
  }
}

/**
 * 從比對結果填入人員複選框
 */
function populateComparePersonCheckboxes(items) {
  const comparePersonCheckboxGroup = document.getElementById('comparePersonCheckboxGroup');
  if (!comparePersonCheckboxGroup) return;
  const nameSet = {};
  items.forEach(function(item) {
    if (item.schedule && item.schedule.name) nameSet[item.schedule.name] = true;
    if (item.attendance && item.attendance.name) nameSet[item.attendance.name] = true;
  });
  const names = Object.keys(nameSet).sort();
  comparePersonCheckboxGroup.innerHTML = '';
  if (names.length > 0) {
    names.forEach(function(n) {
      const label = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = n;
      label.appendChild(cb);
      label.appendChild(document.createTextNode(n));
      comparePersonCheckboxGroup.appendChild(label);
    });
  } else {
    comparePersonCheckboxGroup.innerHTML = '<span class="person-placeholder">此範圍無人員資料</span>';
  }
}

/**
 * 渲染比對結果卡片
 */
function renderCompareResults(items) {
  const compareList = document.getElementById('compareList');
  if (!compareList) return;

  if (!items || items.length === 0) {
    compareList.innerHTML = '<p class="person-placeholder">無比對資料</p>';
    return;
  }

  compareList.innerHTML = items.map(function(item, idx) {
    const s = item.schedule || null;
    const a = item.attendance || null;
    const corr = item.correction || null;
    const displayName = (s && s.name) || (a && a.name) || '—';
    const empAccount = (s && s.empAccount) || (a && a.empAccount) || '';
    const branch = (s && s.branch) || (a && a.branch) || '';
    const date = (s && s.date) || (a && a.date) || '';
    const scheduleStart = s ? (s.startTime || '—') : '—';
    const scheduleEnd = s ? (s.endTime || '—') : '—';
    const scheduleHours = s ? (s.hours || '—') : '—';
    const attendanceStart = a ? (a.startTime || '—') : '—';
    const attendanceEnd = a ? (a.endTime || '—') : '—';
    const attendanceHours = a ? (a.hours || '—') : '—';
    const attendanceStatus = a ? (a.status || '—') : '—';

    const correctedStart = corr ? corr.correctedStart : '';
    const correctedEnd = corr ? corr.correctedEnd : '';
    const isCorrected = !!(corr && correctedStart && correctedEnd);

    const payload = JSON.stringify({
      branch: branch,
      empAccount: empAccount,
      name: displayName,
      date: date,
      scheduleStart: s ? s.startTime : '',
      scheduleEnd: s ? s.endTime : '',
      scheduleHours: s ? s.hours : '',
      attendanceStart: a ? a.startTime : '',
      attendanceEnd: a ? a.endTime : '',
      attendanceHours: a ? a.hours : '',
      attendanceStatus: a ? a.status : ''
    });

    const cardId = 'compare-card-' + idx;

    return (
      '<div class="compare-card' + (isCorrected ? ' corrected' : '') + '" id="' + cardId + '" data-payload="' + escapeHtmlAttr(payload) + '">' +
        '<div class="result-row"><span class="result-label">分店</span><span class="result-value">' + escapeHtml(branch) + '</span></div>' +
        '<div class="result-row"><span class="result-label">員工帳號</span><span class="result-value">' + escapeHtml(empAccount) + '</span></div>' +
        '<div class="result-row"><span class="result-label">姓名</span><span class="result-value">' + escapeHtml(displayName) + '</span></div>' +
        '<div class="result-row"><span class="result-label">日期</span><span class="result-value">' + escapeHtml(date) + '</span></div>' +
        '<div class="result-row"><span class="result-label">班表 上班/下班/時數</span><span class="result-value">' + escapeHtml(scheduleStart) + ' / ' + escapeHtml(scheduleEnd) + ' / ' + escapeHtml(scheduleHours) + '</span></div>' +
        '<div class="result-row"><span class="result-label">打卡 上班/下班/時數/狀態</span><span class="result-value">' + escapeHtml(attendanceStart) + ' / ' + escapeHtml(attendanceEnd) + ' / ' + escapeHtml(attendanceHours) + ' / ' + escapeHtml(attendanceStatus) + '</span></div>' +
        '<div class="compare-card-actions">' +
          '<label><span class="result-label">校正上班</span><input type="text" class="corrected-start-input schedule-date-input" placeholder="HH:mm" value="' + escapeHtmlAttr(correctedStart) + '" ' + (isCorrected ? 'readonly' : '') + '></label>' +
          '<label><span class="result-label">校正下班</span><input type="text" class="corrected-end-input schedule-date-input" placeholder="HH:mm" value="' + escapeHtmlAttr(correctedEnd) + '" ' + (isCorrected ? 'readonly' : '') + '></label>' +
          (isCorrected
            ? '<span class="compare-card-badge">已校正</span><button type="button" class="person-btn edit-correction-btn">編輯</button>'
            : '<button type="button" class="load-schedule-btn submit-correction-btn">校正送出</button>') +
        '</div>' +
      '</div>'
    );
  }).join('');

  compareList.querySelectorAll('.submit-correction-btn').forEach(function(btn) {
    btn.addEventListener('click', handleSubmitCorrectionClick);
  });
  compareList.querySelectorAll('.edit-correction-btn').forEach(function(btn) {
    btn.addEventListener('click', handleEditCorrectionClick);
  });
}

function escapeHtml(s) {
  if (s === undefined || s === null) return '';
  const t = String(s);
  return t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function escapeHtmlAttr(s) {
  if (s === undefined || s === null) return '';
  return escapeHtml(s).replace(/'/g, '&#39;');
}

/**
 * 處理校正送出按鈕點擊
 */
function handleSubmitCorrectionClick(e) {
  const btn = e.target;
  const card = btn.closest('.compare-card');
  if (!card) return;
  const payloadStr = card.getAttribute('data-payload');
  if (!payloadStr) return;
  let payload;
  try {
    payload = JSON.parse(payloadStr);
  } catch (err) {
    showAlert('error', '資料格式錯誤');
    return;
  }
  const correctedStartInput = card.querySelector('.corrected-start-input');
  const correctedEndInput = card.querySelector('.corrected-end-input');
  const correctedStart = correctedStartInput ? correctedStartInput.value.trim() : '';
  const correctedEnd = correctedEndInput ? correctedEndInput.value.trim() : '';
  if (!correctedStart || !correctedEnd) {
    showAlert('error', '請填寫校正上班時間與校正下班時間');
    return;
  }
  payload.correctedStart = correctedStart;
  payload.correctedEnd = correctedEnd;
  doSubmitCorrection(payload);
}

/**
 * 處理編輯按鈕點擊（已校正狀態下切換為可編輯）
 */
function handleEditCorrectionClick(e) {
  const btn = e.target;
  const card = btn.closest('.compare-card');
  if (!card) return;
  const correctedStartInput = card.querySelector('.corrected-start-input');
  const correctedEndInput = card.querySelector('.corrected-end-input');
  if (correctedStartInput) correctedStartInput.removeAttribute('readonly');
  if (correctedEndInput) correctedEndInput.removeAttribute('readonly');
  const badge = card.querySelector('.compare-card-badge');
  if (badge) badge.remove();
  const newBtn = document.createElement('button');
  newBtn.type = 'button';
  newBtn.className = 'load-schedule-btn submit-correction-btn';
  newBtn.textContent = '校正送出';
  newBtn.addEventListener('click', handleSubmitCorrectionClick);
  btn.replaceWith(newBtn);
}

/**
 * 送出校正到 API
 */
async function doSubmitCorrection(payload) {
  try {
    const response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({
        action: 'submitCorrection',
        branch: payload.branch,
        empAccount: payload.empAccount,
        name: payload.name,
        date: payload.date,
        scheduleStart: payload.scheduleStart,
        scheduleEnd: payload.scheduleEnd,
        scheduleHours: payload.scheduleHours,
        attendanceStart: payload.attendanceStart,
        attendanceEnd: payload.attendanceEnd,
        attendanceHours: payload.attendanceHours,
        attendanceStatus: payload.attendanceStatus,
        correctedStart: payload.correctedStart,
        correctedEnd: payload.correctedEnd
      })
    });
    const result = await response.json();
    if (!response.ok || !result.success) {
      throw new Error(result.error || '校正送出失敗');
    }
    showAlert('success', '校正紀錄已送出');
    const loadCompareBtn = document.getElementById('loadCompareBtn');
    if (loadCompareBtn) loadCompareBtn.click();
  } catch (error) {
    showAlert('error', '校正送出失敗：' + error.message);
  }
}
