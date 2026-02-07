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

// 上傳類型切換（更新工作表名稱說明）
document.querySelectorAll('input[name="uploadType"]').forEach(function(radio) {
  radio.addEventListener('change', function() {
    if (sheetNameHint) {
      sheetNameHint.textContent = this.value === 'schedule'
        ? '請輸入要處理的 Excel 工作表名稱（例如：11501、11502）'
        : '請輸入要處理的 Excel 工作表名稱（例如：打卡紀錄、Sheet1）';
    }
    if (sheetNameInput && !sheetNameInput.value) {
      sheetNameInput.placeholder = this.value === 'schedule' ? '例如：11501' : '例如：打卡紀錄';
    }
  });
});

// 載入國安班表按鈕
if (loadScheduleBtn) loadScheduleBtn.addEventListener('click', handleLoadSchedule);

// 日期篩選模式切換
document.querySelectorAll('input[name="dateFilterMode"]').forEach(radio => {
  if (radio) radio.addEventListener('change', toggleDateFilterMode);
});
if (datePicker) datePicker.addEventListener('change', function() {});

// 人員篩選按鈕
if (selectAllPersonsBtn) selectAllPersonsBtn.addEventListener('click', selectAllPersons);
if (clearAllPersonsBtn) clearAllPersonsBtn.addEventListener('click', clearAllPersons);

// 頁面載入時初始化
document.addEventListener('DOMContentLoaded', function() {
  if (yearSelect && monthSelect) initScheduleSelectors();
  toggleDateFilterMode();
});

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
 * 格式化工作時數為分鐘顯示（後端為小時，7.5 → 450分）
 */
function formatMinutes(val) {
  if (val === undefined || val === null || String(val).trim() === '') return '—';
  var s = String(val).trim();
  var m = s.match(/[\d.]+/);
  if (!m) return s;
  var num = parseFloat(m[0]);
  if (isNaN(num)) return s;
  var minutes = Math.round(num * 60);
  return minutes + '分';
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
  
  // 取得並驗證工作表名稱
  const sheetName = sheetNameInput.value.trim();
  if (!sheetName) {
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
    const payload = {
      action: 'upload',
      uploadType: uploadType ? uploadType.value : 'schedule',
      fileName: selectedFile.name,
      fileData: base64Data,
      targetSheetName: sheetName,
      targetGoogleSheetName: CONFIG.TARGET_GOOGLE_SHEET_NAME,
      targetGoogleSheetTab: uploadType && uploadType.value === 'attendance' ? '打卡紀錄' : CONFIG.TARGET_GOOGLE_SHEET_TAB
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

  const summaryItems = [
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

  resultList.innerHTML = recordList.map(({ row, isDuplicate }) => {
    const [
      name,
      date,
      start,
      end,
      hours,
      shift
    ] = row;
    const duplicateBadge = isDuplicate ? '<span class="result-card-duplicate" title="此筆為重複，已略過寫入">重複</span>' : '';
    return `
      <div class="result-card ${isDuplicate ? 'result-card--duplicate' : ''}">
        ${duplicateBadge}
        <div class="result-row"><span class="result-label">姓名</span><span class="result-value">${name || '—'}</span></div>
        <div class="result-row"><span class="result-label">日期</span><span class="result-value">${date || '—'}</span></div>
        <div class="result-row"><span class="result-label">班別</span><span class="result-value">${shift || '—'}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">分鐘</span><span class="result-value">${formatMinutes(hours)}</span></div>
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
 * 載入國安班表（依年月/日期 + 人員篩選，AND 關係）
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
  let url = `${CONFIG.GAS_URL}?action=loadSchedule`;
  if (yearMonth) url += `&yearMonth=${encodeURIComponent(yearMonth)}`;
  if (dateParam) url += `&date=${encodeURIComponent(dateParam)}`;
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
    showAlert('error', '載入國安班表失敗：' + error.message);
  } finally {
    loadScheduleBtn.disabled = false;
    loadScheduleBtn.textContent = '載入';
  }
}

/**
 * 顯示國安班表查詢結果
 */
function renderScheduleResults(result) {
  const details = result.details || {};
  const records = Array.isArray(result.records) ? result.records : [];
  const names = Array.isArray(details.names) ? details.names : [];

  scheduleSummary.innerHTML = `
    <div class="summary-item">
      <div class="summary-label">日期範圍</div>
      <div class="summary-value">${details.date ? details.date.replace(/-/g, '/') : (details.yearMonth ? details.yearMonth.substring(0,4) + '/' + details.yearMonth.substring(4,6) : '—')}</div>
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
    const [name, date, start, end, hours, shift] = row;
    return `
      <div class="result-card">
        <div class="result-row"><span class="result-label">姓名</span><span class="result-value">${name || '—'}</span></div>
        <div class="result-row"><span class="result-label">日期</span><span class="result-value">${date || '—'}</span></div>
        <div class="result-row"><span class="result-label">班別</span><span class="result-value">${shift || '—'}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">分鐘</span><span class="result-value">${formatMinutes(hours)}</span></div>
      </div>
    `;
  }).join('');

  scheduleResultSection.classList.add('show');
}
