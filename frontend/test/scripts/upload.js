/**
 * 薪資計算系統 - 檔案上傳邏輯
 */

let selectedFile = null;

// DOM 元素
const fileInput = document.getElementById('fileInput');
const fileInfo = document.getElementById('fileInfo');
const fileSelectBtn = document.getElementById('fileSelectBtn');
const sheetNameInput = document.getElementById('sheetName');
const sheetNameHint = document.getElementById('sheetNameHint');
const sheetNameGroup = document.getElementById('sheetNameGroup');
const submitBtn = document.getElementById('submitBtn');
const progressContainer = document.getElementById('progressContainer');
const progressFill = document.getElementById('progressFill');
const progressText = document.getElementById('progressText');
const alertModal = document.getElementById('alertModal');
const confirmModal = document.getElementById('confirmModal');
const resultSection = document.getElementById('resultSection');
const resultSummary = document.getElementById('resultSummary');
const resultList = document.getElementById('resultList');
const queryYearMonthInput = document.getElementById('queryYearMonthInput');
const queryStartDate = document.getElementById('queryStartDate');
const queryEndDate = document.getElementById('queryEndDate');
const loadScheduleBtn = document.getElementById('loadScheduleBtn');
const scheduleResultSection = document.getElementById('scheduleResultSection');
const scheduleSummary = document.getElementById('scheduleSummary');
const scheduleList = document.getElementById('scheduleList');
const branchSelect = document.getElementById('branchSelect');
const branchGroup = document.getElementById('branchGroup');

// 查詢/比對共用人員名單（載入後填入 personCheckboxGroup）
var __personnelNames = [];


// 檔案選擇事件
fileInput.addEventListener('change', handleFileSelect);

// 選擇檔案按鈕 → 觸發 file input
if (fileSelectBtn) fileSelectBtn.addEventListener('click', function() { fileInput.click(); });

// 提交按鈕事件
submitBtn.addEventListener('click', handleSubmit);

// 依上傳類型設定 file input accept
// 打卡：Line / 部分 WebView 對 accept 支援差，改不限制類型，改由選檔後 JS 驗證副檔名為 .csv
function setFileInputAccept() {
  if (!fileInput) return;
  var uploadType = document.querySelector('input[name="uploadType"]:checked');
  var isAttendance = uploadType && uploadType.value === 'attendance';
  fileInput.accept = isAttendance ? '' : '.xlsx,.xls';
}

// 上傳類型切換（分店、工作表區塊顯示、file accept）
document.querySelectorAll('input[name="uploadType"]').forEach(function(radio) {
  radio.addEventListener('change', function() {
    var isSchedule = this.value === 'schedule';
    var isAttendance = this.value === 'attendance';
    setFileInputAccept();
    if (sheetNameHint) sheetNameHint.textContent = isSchedule ? '選檔後自動帶入工作表清單' : '打卡上傳 CSV 不需選擇工作表';
    if (sheetNameGroup) sheetNameGroup.style.display = isSchedule ? 'block' : 'none';
    if (branchGroup) branchGroup.style.display = (isSchedule || isAttendance) ? 'block' : 'none';
    if (selectedFile) {
      var ext = '.' + selectedFile.name.split('.').pop().toLowerCase();
      var ok = isAttendance ? ext === '.csv' : CONFIG.ALLOWED_FILE_TYPES.includes(ext);
      if (!ok) {
        selectedFile = null;
        fileInput.value = '';
        if (fileInfo) fileInfo.textContent = '';
        submitBtn.classList.remove('show');
        resetSheetSelect();
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

// 查詢類型切換（班表／打卡）：若已有載入結果則清除，避免標題與資料不符
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
    if (scheduleResultSection && scheduleResultSection.classList.contains('show')) {
      if (scheduleSummary) scheduleSummary.innerHTML = '';
      if (scheduleList) scheduleList.innerHTML = '';
      scheduleResultSection.classList.remove('show');
    }
  });
});
// 查詢/比對共用：選擇分店時載入人員
const queryBranchSelect = document.getElementById('queryBranchSelect');
if (queryBranchSelect) queryBranchSelect.addEventListener('change', handleQueryBranchChange);

// 人員全選／清除
var selectAllPersonsBtn = document.getElementById('selectAllPersonsBtn');
var clearAllPersonsBtn = document.getElementById('clearAllPersonsBtn');
if (selectAllPersonsBtn) selectAllPersonsBtn.addEventListener('click', selectAllPersons);
if (clearAllPersonsBtn) clearAllPersonsBtn.addEventListener('click', clearAllPersons);

// 載入比對按鈕
const loadCompareBtn = document.getElementById('loadCompareBtn');
if (loadCompareBtn) loadCompareBtn.addEventListener('click', handleLoadCompare);

// 結果區收合鈕
document.addEventListener('click', function(e) {
  var btn = e.target && e.target.closest && e.target.closest('.collapse-toggle-btn');
  if (!btn) return;
  var targetId = btn.getAttribute('data-target');
  if (!targetId) return;
  var body = document.getElementById(targetId);
  if (!body) return;
  body.classList.toggle('collapsed');
  btn.classList.toggle('collapsed');
  btn.textContent = body.classList.contains('collapsed') ? '▶' : '▼';
});

// 結果區收合鈕：點擊 overlay 或關閉按鈕關閉 modal
if (alertModal) {
  alertModal.addEventListener('click', function(e) {
    if (e.target.classList.contains('alert-modal-overlay') || e.target.classList.contains('alert-modal-close')) {
      hideAlert();
    }
  });
}

var confirmModalCallback = null;
if (confirmModal) {
  confirmModal.addEventListener('click', function(e) {
    if (e.target.classList.contains('confirm-modal-overlay') || e.target.classList.contains('confirm-modal-cancel')) {
      hideConfirm();
      confirmModalCallback = null;
    } else if (e.target.classList.contains('confirm-modal-ok')) {
      var cb = confirmModalCallback;
      hideConfirm();
      confirmModalCallback = null;
      if (cb) cb();
    }
  });
}

// Tab 切換：一次只顯示一個功能區塊；查詢/比對時顯示共用條件區塊
function initTabNav() {
  var nav = document.querySelector('.tab-nav');
  var sharedBlock = document.getElementById('sharedConditionBlock');
  if (!nav) return;
  function updateSharedConditionVisibility(tabId) {
    if (sharedBlock) sharedBlock.style.display = (tabId === 'query' || tabId === 'compare') ? 'block' : 'none';
  }
  nav.addEventListener('click', function(e) {
    var btn = e.target && e.target.closest && e.target.closest('.tab-nav-btn');
    if (!btn || btn.classList.contains('active')) return;
    var tabId = btn.getAttribute('data-tab');
    var paneId = 'tab-pane-' + tabId;
    var pane = document.getElementById(paneId);
    if (!pane) return;
    document.querySelectorAll('.tab-nav-btn').forEach(function(b) {
      b.classList.remove('active');
      b.setAttribute('aria-selected', 'false');
    });
    document.querySelectorAll('.tab-pane').forEach(function(p) {
      p.classList.remove('active');
      p.setAttribute('hidden', '');
    });
    btn.classList.add('active');
    btn.setAttribute('aria-selected', 'true');
    pane.classList.add('active');
    pane.removeAttribute('hidden');
    updateSharedConditionVisibility(tabId);
    if (tabId === 'query' || tabId === 'compare') loadQueryPersonnel();
  });
  updateSharedConditionVisibility('upload');
}

function selectAllPersons() {
  var group = document.getElementById('personCheckboxGroup');
  if (group) group.querySelectorAll('input[type="checkbox"]').forEach(function(cb) { cb.checked = true; });
}

function clearAllPersons() {
  var group = document.getElementById('personCheckboxGroup');
  if (group) group.querySelectorAll('input[type="checkbox"]').forEach(function(cb) { cb.checked = false; });
}

// 頁面載入時初始化
document.addEventListener('DOMContentLoaded', function() {
  initTabNav();
  initYearMonthSelects();
  toggleDateFilterMode();
  loadBranches();
  // 共用人員：年月／日期變更時重新載入人員
  var qymSel = document.getElementById('queryYearMonthSelect');
  var qymInp = document.getElementById('queryYearMonthInput');
  if (qymSel) qymSel.addEventListener('change', loadQueryPersonnel);
  if (qymInp) { qymInp.addEventListener('input', loadQueryPersonnel); qymInp.addEventListener('change', loadQueryPersonnel); }
  if (queryStartDate) queryStartDate.addEventListener('change', loadQueryPersonnel);
  if (queryEndDate) queryEndDate.addEventListener('change', loadQueryPersonnel);
  // 初始顯示分店、工作表區塊與 file accept（依上傳類型）
  var mode = document.querySelector('input[name="uploadType"]:checked');
  if (branchGroup) branchGroup.style.display = mode && (mode.value === 'schedule' || mode.value === 'attendance') ? 'block' : 'none';
  if (sheetNameGroup) sheetNameGroup.style.display = mode && mode.value === 'schedule' ? 'block' : 'none';
  setFileInputAccept();
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
    showLoadingOverlay();
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
      queryBranchEl.innerHTML = '<option value="">請選擇分店</option>';
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
    loadQueryPersonnel();
  } catch (error) {
    console.error('載入分店清單失敗:', error);
    if (branchEl) branchEl.innerHTML = '<option value="">載入失敗，請重整頁面</option>';
    if (queryBranchEl) queryBranchEl.innerHTML = '<option value="">載入失敗</option>';
    if (compareBranchEl) compareBranchEl.innerHTML = '<option value="">載入失敗</option>';
  } finally {
    hideLoadingOverlay();
  }
}

/**
 * 處理檔案選擇
 */
function handleFileSelect(e) {
  var file = e.target.files[0];
  if (file) validateAndDisplayFile(file);
}

/**
 * 重設工作表下拉（無選檔或打卡時）
 */
function resetSheetSelect() {
  if (!sheetNameInput) return;
  sheetNameInput.innerHTML = '<option value="">請先選擇檔案</option>';
  sheetNameInput.value = '';
  sheetNameInput.disabled = true;
}

/**
 * 從 Excel 檔案解析工作表名稱並填入下拉
 */
function parseAndFillSheetNames(file) {
  return new Promise(function(resolve, reject) {
    if (!sheetNameInput) return resolve();
    if (typeof XLSX === 'undefined') {
      sheetNameInput.innerHTML = '<option value="">需載入 xlsx 套件</option>';
      sheetNameInput.disabled = true;
      return resolve();
    }
    var reader = new FileReader();
    reader.onload = function(e) {
      try {
        var ab = e.target.result;
        var workbook = XLSX.read(ab, { type: 'arraybuffer', bookSheets: true });
        var names = workbook.SheetNames || [];
        sheetNameInput.innerHTML = '';
        var opt0 = document.createElement('option');
        opt0.value = '';
        opt0.textContent = names.length ? '請選擇工作表' : '此檔案無工作表';
        sheetNameInput.appendChild(opt0);
        names.forEach(function(n) {
          var opt = document.createElement('option');
          opt.value = n;
          opt.textContent = n;
          sheetNameInput.appendChild(opt);
        });
        sheetNameInput.disabled = names.length === 0;
        if (names.length === 1) sheetNameInput.value = names[0];
      } catch (err) {
        console.error('parse sheet names:', err);
        sheetNameInput.innerHTML = '<option value="">解析失敗</option>';
        sheetNameInput.disabled = true;
      }
      resolve();
    };
    reader.onerror = function() { reject(new Error('讀取檔案失敗')); };
    reader.readAsArrayBuffer(file);
  });
}

/**
 * 驗證並顯示檔案資訊，班表時並解析工作表名稱填入下拉
 */
function validateAndDisplayFile(file) {
  var fileExtension = '.' + file.name.split('.').pop().toLowerCase();
  var uploadType = document.querySelector('input[name="uploadType"]:checked');
  var isAttendance = uploadType && uploadType.value === 'attendance';
  var allowedTypes = isAttendance ? ['.csv'] : CONFIG.ALLOWED_FILE_TYPES;
  if (!allowedTypes.includes(fileExtension)) {
    showAlert('error', isAttendance ? '打卡上傳請使用 .csv 檔案' : '班表上傳請使用 ' + CONFIG.ALLOWED_FILE_TYPES.join('、') + ' 檔案');
    return;
  }
  if (file.size > CONFIG.MAX_FILE_SIZE) {
    var maxSizeMB = CONFIG.MAX_FILE_SIZE / (1024 * 1024);
    showAlert('error', '檔案過大。最大允許 ' + maxSizeMB + 'MB。');
    return;
  }
  selectedFile = file;
  if (fileInfo) fileInfo.textContent = file.name + ' (' + formatFileSize(file.size) + ')';
  submitBtn.classList.add('show');
  submitBtn.disabled = false;
  submitBtn.textContent = '開始上傳並處理';
  progressContainer.classList.remove('show');
  clearResults();
  hideAlert();
  if (isAttendance) {
    resetSheetSelect();
    return;
  }
  if (fileExtension === '.xlsx' || fileExtension === '.xls') {
    parseAndFillSheetNames(file);
  } else {
    sheetNameInput.disabled = false;
  }
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

  // 班表需選擇工作表；打卡不需（CSV 為整檔）
  var sheetName = sheetNameInput && sheetNameInput.value ? sheetNameInput.value.trim() : '';
  if (uploadType && uploadType.value === 'schedule' && !sheetName) {
    showAlert('error', '請選擇 Excel 工作表名稱');
    return;
  }
  
  // 禁用提交按鈕和輸入欄位
  submitBtn.disabled = true;
  submitBtn.textContent = '處理中...';
  sheetNameInput.disabled = true;
  
  // 顯示進度條與 overlay
  progressContainer.classList.add('show');
  showLoadingOverlay();
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
    showAlert('success', '✅ 解析完成，結果如下');
    
  } catch (error) {
    console.error('上傳錯誤:', error);
    showAlert('error', `上傳失敗: ${error.message}`);
    submitBtn.disabled = false;
    submitBtn.textContent = '開始上傳並處理';
    sheetNameInput.disabled = false;
    progressContainer.classList.remove('show');
  } finally {
    hideLoadingOverlay();
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
 * 顯示提示訊息（跳出視窗）
 * @param {string} type - success | error | warning
 * @param {string} message - 訊息內容
 * @param {Object} options - 保留供擴充，目前不捲動畫面以維持原位置
 */
var alertAutoCloseTimer = null;

function showAlert(type, message, options) {
  if (!alertModal) return;
  if (alertAutoCloseTimer) {
    clearTimeout(alertAutoCloseTimer);
    alertAutoCloseTimer = null;
  }
  var content = alertModal.querySelector('.alert-modal-content');
  var msgEl = alertModal.querySelector('.alert-modal-message');
  if (content) {
    content.className = 'alert-modal-content alert-' + type;
  }
  if (msgEl) msgEl.textContent = message;
  alertModal.classList.add('show');
  if (type === 'success') {
    alertAutoCloseTimer = setTimeout(function() {
      hideAlert();
      alertAutoCloseTimer = null;
    }, 3000);
  }
}

/**
 * 顯示執行中 overlay
 */
function showLoadingOverlay() {
  var el = document.getElementById('loadingOverlay');
  if (el) { el.classList.add('show'); el.setAttribute('aria-hidden', 'false'); }
}

/**
 * 隱藏執行中 overlay
 */
function hideLoadingOverlay() {
  var el = document.getElementById('loadingOverlay');
  if (el) { el.classList.remove('show'); el.setAttribute('aria-hidden', 'true'); }
}

/**
 * 顯示確認視窗（自訂 modal，不顯示 domain）
 * @param {string} message - 訊息內容
 * @param {function} onConfirm - 按下「確定」時執行
 */
function showConfirm(message, onConfirm) {
  if (!confirmModal) return;
  var msgEl = confirmModal.querySelector('.confirm-modal-message');
  if (msgEl) msgEl.textContent = message;
  confirmModalCallback = onConfirm;
  confirmModal.classList.add('show');
}

/**
 * 隱藏確認視窗
 */
function hideConfirm() {
  if (confirmModal) confirmModal.classList.remove('show');
}

/**
 * 隱藏提示訊息
 */
function hideAlert() {
  if (alertAutoCloseTimer) {
    clearTimeout(alertAutoCloseTimer);
    alertAutoCloseTimer = null;
  }
  if (alertModal) alertModal.classList.remove('show');
}

/**
 * 重置表單
 */
function resetForm() {
  selectedFile = null;
  fileInput.value = '';
  if (fileInfo) fileInfo.textContent = '';
  submitBtn.classList.remove('show');
  submitBtn.disabled = false;
  submitBtn.textContent = '開始上傳並處理';
  resetSheetSelect();
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

  const deletedCount = details.deletedCount ?? 0;
  const baseAttendanceItems = [
    { label: '新增筆數', value: details.rowCount ?? '—' },
    ...(deletedCount > 0 ? [{ label: '已覆蓋筆數', value: deletedCount }] : []),
    { label: '略過重複', value: details.skippedCount ?? 0 },
    { label: '原始筆數', value: details.parsedRowCount ?? records.length ?? '—' },
    { label: '處理時間', value: details.processTime ? `${details.processTime}s` : '—' },
    { label: '目標工作表', value: details.targetSheet || '—' }
  ];
  const baseScheduleItems = [
    { label: '新增筆數', value: details.rowCount ?? '—' },
    ...(deletedCount > 0 ? [{ label: '已覆蓋筆數', value: deletedCount }] : []),
    { label: '略過重複', value: details.skippedCount ?? 0 },
    { label: '原始筆數', value: details.parsedRowCount ?? records.length ?? '—' },
    { label: '員工數', value: details.totalEmployees || 0 },
    { label: '班別代碼', value: shiftCodes.length ? shiftCodes.join(', ') : '—' },
    { label: '處理時間', value: details.processTime ? `${details.processTime}s` : '—' },
    { label: '來源工作表', value: details.sourceSheet || '—' },
    { label: '目標工作表', value: details.targetSheet || '—' }
  ];
  const isAttendanceResult = result.columns && result.columns[0] === '分店' && result.columns[1] === '員工編號';
  const summaryItems = isAttendanceResult ? baseAttendanceItems : baseScheduleItems;

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
        <div class="result-row"><span class="result-label">打卡日期</span><span class="result-value">${formatDateWithWeekday(date)}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">分店</span><span class="result-value">${branch || '—'}</span></div>
        <div class="result-row"><span class="result-label">工作時數</span><span class="result-value">${formatHoursWithMinutes(hours, start, end)}</span></div>
        ${status && String(status).trim() ? '<div class="result-row"><span class="result-label">狀態</span><span class="result-value">' + escapeHtml(status) + '</span></div>' : ''}
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
        <div class="result-row"><span class="result-label">日期</span><span class="result-value">${formatDateWithWeekday(date)}</span></div>
        <div class="result-row"><span class="result-label">班別</span><span class="result-value">${shift || '—'}</span></div>
        <div class="result-row"><span class="result-label">分店</span><span class="result-value">${branch || '—'}</span></div>
        <div class="result-row"><span class="result-label">上班</span><span class="result-value">${start || '—'}</span></div>
        <div class="result-row"><span class="result-label">下班</span><span class="result-value">${end || '—'}</span></div>
        <div class="result-row"><span class="result-label">時數</span><span class="result-value">${formatHoursWithMinutes(hours, start, end)}</span></div>
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

/**
 * 產生上月、本月、下月 YYYYMM 選項，預設本月
 */
function getYearMonthOptions() {
  var now = new Date();
  var y = now.getFullYear();
  var m = now.getMonth();
  var items = [];
  for (var i = -1; i <= 1; i++) {
    var d = new Date(y, m + i, 1);
    var ym = d.getFullYear() * 100 + (d.getMonth() + 1);
    var label = i === -1 ? '上月' : (i === 0 ? '本月' : '下月');
    items.push({ label: label, value: String(ym) });
  }
  return items;
}

/**
 * 初始化年月下拉（上月/本月/下月），預設本月
 */
function initYearMonthSelects() {
  var opts = getYearMonthOptions();
  var currentYm = opts[1].value;
  var querySel = document.getElementById('queryYearMonthSelect');
  function fillSelect(sel) {
    if (!sel) return;
    sel.innerHTML = opts.map(function(o) {
      return '<option value="' + o.value + '"' + (o.value === currentYm ? ' selected' : '') + '>' + o.label + ' (' + o.value + ')</option>';
    }).join('');
  }
  fillSelect(querySel);
  if (queryYearMonthInput) queryYearMonthInput.value = currentYm;
  function onSelectChange(sel, input) {
    if (!sel || !input) return;
    sel.addEventListener('change', function() {
      input.value = sel.value;
    });
  }
  onSelectChange(querySel, queryYearMonthInput);
}

function toggleDateFilterMode() {
  const mode = document.querySelector('input[name="dateFilterMode"]:checked');
  const isMonth = mode && mode.value === 'month';
  const dateMonthGroup = document.getElementById('dateMonthGroup');
  const dateRangeGroup = document.getElementById('dateRangeGroup');
  var querySel = document.getElementById('queryYearMonthSelect');
  if (querySel) querySel.disabled = !isMonth;
  if (queryYearMonthInput) queryYearMonthInput.disabled = !isMonth;
  if (queryStartDate) queryStartDate.disabled = isMonth;
  if (queryEndDate) queryEndDate.disabled = isMonth;
  if (dateMonthGroup) dateMonthGroup.classList.toggle('hidden', !isMonth);
  if (dateRangeGroup) dateRangeGroup.classList.toggle('hidden', isMonth);
  loadQueryPersonnel();
}

/**
 * 載入查詢/比對共用人員名單（以該月份／日期區間＋分店的打卡資料為來源）
 * 結果填入 personCheckboxGroup，保留目前勾選狀態
 */
async function loadQueryPersonnel() {
  var personCheckboxGroup = document.getElementById('personCheckboxGroup');
  var selectAllBtn = document.getElementById('selectAllPersonsBtn');
  var clearBtn = document.getElementById('clearAllPersonsBtn');
  if (!personCheckboxGroup) return;
  var branchVal = (document.getElementById('queryBranchSelect') && document.getElementById('queryBranchSelect').value) ? document.getElementById('queryBranchSelect').value.trim() : '';
  var mode = document.querySelector('input[name="dateFilterMode"]:checked');
  var isMonth = mode && mode.value === 'month';
  var yearMonth = '';
  var startDate = '';
  var endDate = '';
  if (isMonth) {
    var sel = document.getElementById('queryYearMonthSelect');
    var inp = document.getElementById('queryYearMonthInput');
    yearMonth = (inp && inp.value && inp.value.trim().match(/^\d{6}$/)) ? inp.value.trim() : (sel && sel.value ? sel.value : '');
  } else {
    var qStart = document.getElementById('queryStartDate');
    var qEnd = document.getElementById('queryEndDate');
    startDate = qStart && qStart.value ? qStart.value.trim() : '';
    endDate = (qEnd && qEnd.value ? qEnd.value.trim() : '') || startDate;
  }
  if (!branchVal) {
    __personnelNames = [];
    personCheckboxGroup.innerHTML = '<span class="person-placeholder">請先選擇分店</span>';
    if (selectAllBtn) selectAllBtn.disabled = true;
    if (clearBtn) clearBtn.disabled = true;
    return;
  }
  if (!yearMonth && !startDate) {
    __personnelNames = [];
    personCheckboxGroup.innerHTML = '<span class="person-placeholder">請選擇月份或日期區間以載入人員</span>';
    if (selectAllBtn) selectAllBtn.disabled = true;
    if (clearBtn) clearBtn.disabled = true;
    return;
  }
  var currentChecked = getSelectedPersonNames();
  personCheckboxGroup.innerHTML = '<span class="person-placeholder">載入中...</span>';
  if (selectAllBtn) selectAllBtn.disabled = true;
  if (clearBtn) clearBtn.disabled = true;
  try {
    showLoadingOverlay();
    var params = 'action=getPersonnelFromSchedule&branch=' + encodeURIComponent(branchVal);
    if (yearMonth) params += '&yearMonth=' + encodeURIComponent(yearMonth); else { params += '&startDate=' + encodeURIComponent(startDate); params += '&endDate=' + encodeURIComponent(endDate); }
    var response = await fetch(CONFIG.GAS_URL + '?' + params, { method: 'GET', mode: 'cors' });
    var result = await response.json();
    __personnelNames = (result.success && Array.isArray(result.names)) ? result.names : [];
    renderQueryPersonCheckboxes(__personnelNames, { checked: currentChecked }, true);
  } catch (error) {
    console.error('載入人員失敗:', error);
    __personnelNames = [];
    personCheckboxGroup.innerHTML = '<span class="person-placeholder">載入失敗，請重整頁面</span>';
    if (selectAllBtn) selectAllBtn.disabled = true;
    if (clearBtn) clearBtn.disabled = true;
  } finally {
    hideLoadingOverlay();
  }
}

function renderQueryPersonCheckboxes(names, opts, fromSchedule) {
  var personCheckboxGroup = document.getElementById('personCheckboxGroup');
  var selectAllPersonsBtn = document.getElementById('selectAllPersonsBtn');
  var clearAllPersonsBtn = document.getElementById('clearAllPersonsBtn');
  if (!personCheckboxGroup) return;
  var checkedSet = {};
  if (opts && Array.isArray(opts.checked)) opts.checked.forEach(function(n) { checkedSet[n] = true; });
  personCheckboxGroup.innerHTML = '';
  if (names.length > 0) {
    names.forEach(function(n) {
      var label = document.createElement('label');
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = n;
      if (checkedSet[n]) cb.checked = true;
      label.appendChild(cb);
      label.appendChild(document.createTextNode(n));
      personCheckboxGroup.appendChild(label);
    });
    if (selectAllPersonsBtn) selectAllPersonsBtn.disabled = false;
    if (clearAllPersonsBtn) clearAllPersonsBtn.disabled = false;
  } else {
    personCheckboxGroup.innerHTML = '<span class="person-placeholder">' + (fromSchedule ? '此月份打卡無人員' : '此分店無人員資料') + '</span>';
    if (selectAllPersonsBtn) selectAllPersonsBtn.disabled = true;
    if (clearAllPersonsBtn) clearAllPersonsBtn.disabled = true;
  }
}

function handleQueryBranchChange() {
  loadQueryPersonnel();
}

function toggleCompareDateMode() {
  // 查詢與比對共用條件區塊，此處不再需要
}


function getSelectedPersonNames() {
  var group = document.getElementById('personCheckboxGroup');
  if (!group) return [];
  return Array.prototype.slice.call(group.querySelectorAll('input[type="checkbox"]:checked')).map(function(cb) { return cb.value; }).filter(Boolean);
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
 * 載入班表（依月份/日期區間、分店必選、人員篩選，與比對區一致）
 */
async function handleLoadSchedule() {
  const mode = document.querySelector('input[name="dateFilterMode"]:checked');
  const isMonth = mode && mode.value === 'month';
  let yearMonth = '';
  let startDate = '';
  let endDate = '';

  if (isMonth) {
    const ym = queryYearMonthInput && queryYearMonthInput.value.trim().match(/^\d{6}$/)
      ? queryYearMonthInput.value.trim()
      : '';
    if (!ym) {
      showAlert('error', '請輸入年月（例如 202601）');
      return;
    }
    yearMonth = ym;
  } else {
    startDate = queryStartDate && queryStartDate.value ? queryStartDate.value.trim() : '';
    endDate = queryEndDate && queryEndDate.value ? queryEndDate.value.trim() : startDate;
    if (!startDate || startDate.length !== 10) {
      showAlert('error', '請選擇日期區間（開始日期）');
      return;
    }
  }

  const queryBranchEl = document.getElementById('queryBranchSelect');
  const branchVal = queryBranchEl && queryBranchEl.value ? queryBranchEl.value.trim() : '';
  if (!branchVal) {
    showAlert('error', '請選擇分店');
    return;
  }

  loadScheduleBtn.disabled = true;
  loadScheduleBtn.textContent = '載入中...';
  hideAlert();
  scheduleResultSection.classList.remove('show');
  showLoadingOverlay();

  const names = getSelectedPersonNames();
  let url = `${CONFIG.GAS_URL}?action=loadSchedule&branch=${encodeURIComponent(branchVal)}`;
  if (yearMonth) url += `&yearMonth=${encodeURIComponent(yearMonth)}`;
  if (startDate) url += `&startDate=${encodeURIComponent(startDate)}`;
  if (endDate) url += `&endDate=${encodeURIComponent(endDate)}`;
  if (names.length > 0) url += `&names=${encodeURIComponent(names.join(','))}`;

  try {
    const response = await fetch(url, { method: 'GET', mode: 'cors' });
    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.error || '載入失敗');
    }

    renderScheduleResults(result);
    mergeQueryPersonFromDetails(result.details);
  } catch (error) {
    showAlert('error', '載入班表失敗：' + error.message);
  } finally {
    loadScheduleBtn.disabled = false;
    loadScheduleBtn.textContent = '載入';
    hideLoadingOverlay();
  }
}

/**
 * 載入打卡（依月份/日期區間、分店必選、人員篩選，與比對區一致）
 */
async function handleLoadAttendance() {
  const mode = document.querySelector('input[name="dateFilterMode"]:checked');
  const isMonth = mode && mode.value === 'month';
  let yearMonth = '';
  let startDate = '';
  let endDate = '';

  if (isMonth) {
    const ym = queryYearMonthInput && queryYearMonthInput.value.trim().match(/^\d{6}$/)
      ? queryYearMonthInput.value.trim()
      : '';
    if (!ym) {
      showAlert('error', '請輸入年月（例如 202601）');
      return;
    }
    yearMonth = ym;
  } else {
    startDate = queryStartDate && queryStartDate.value ? queryStartDate.value.trim() : '';
    endDate = queryEndDate && queryEndDate.value ? queryEndDate.value.trim() : startDate;
    if (!startDate || startDate.length !== 10) {
      showAlert('error', '請選擇日期區間（開始日期）');
      return;
    }
  }

  const queryBranchEl = document.getElementById('queryBranchSelect');
  const branchVal = queryBranchEl && queryBranchEl.value ? queryBranchEl.value.trim() : '';
  if (!branchVal) {
    showAlert('error', '請選擇分店');
    return;
  }

  loadScheduleBtn.disabled = true;
  loadScheduleBtn.textContent = '載入中...';
  hideAlert();
  scheduleResultSection.classList.remove('show');
  showLoadingOverlay();

  const names = getSelectedPersonNames();
  let url = `${CONFIG.GAS_URL}?action=loadAttendance&branch=${encodeURIComponent(branchVal)}`;
  if (yearMonth) url += `&yearMonth=${encodeURIComponent(yearMonth)}`;
  if (startDate) url += `&startDate=${encodeURIComponent(startDate)}`;
  if (endDate) url += `&endDate=${encodeURIComponent(endDate)}`;
  if (names.length > 0) url += `&names=${encodeURIComponent(names.join(','))}`;

  try {
    const response = await fetch(url, { method: 'GET', mode: 'cors' });
    const result = await response.json();

    if (!response.ok || !result.success) {
      throw new Error(result.error || '載入失敗');
    }

    renderAttendanceResults(result);
    mergeQueryPersonFromDetails(result.details);
  } catch (error) {
    showAlert('error', '載入打卡失敗：' + error.message);
  } finally {
    loadScheduleBtn.disabled = false;
    loadScheduleBtn.textContent = '載入';
    hideLoadingOverlay();
  }
}

/**
 * 從查詢結果 details.names 合併到人員複選框（與比對區一致）
 */
function mergeQueryPersonFromDetails(details) {
  if (!details || !Array.isArray(details.names)) return;
  var group = document.getElementById('personCheckboxGroup');
  if (!group) return;
  var existingNames = {};
  group.querySelectorAll('input[type="checkbox"]').forEach(function(cb) { if (cb.value) existingNames[cb.value] = true; });
  details.names.forEach(function(n) { if (n) existingNames[n] = true; });
  var names = Object.keys(existingNames).sort();
  var checkedNames = [];
  group.querySelectorAll('input[type="checkbox"]:checked').forEach(function(cb) { if (cb.value) checkedNames.push(cb.value); });
  renderQueryPersonCheckboxes(names, { checked: checkedNames });
}

/**
 * 顯示班表查詢結果
 */
function renderScheduleResults(result) {
  const details = result.details || {};
  const records = Array.isArray(result.records) ? result.records : [];
  const branchLabel = details.branch ? details.branch : '—';
  const dateRangeLabel = details.startDate
    ? (details.startDate.replace(/-/g, '/') + (details.endDate && details.endDate !== details.startDate ? ' ~ ' + details.endDate.replace(/-/g, '/') : ''))
    : (details.date ? details.date.replace(/-/g, '/') : (details.yearMonth ? details.yearMonth.substring(0,4) + '/' + details.yearMonth.substring(4,6) : '—'));
  scheduleSummary.innerHTML = `
    <div class="summary-item">
      <div class="summary-label">日期範圍</div>
      <div class="summary-value">${dateRangeLabel}</div>
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

  scheduleList.innerHTML = records.map(function(row, idx) {
    var name = row[0];
    var date = row[1];
    var start = row[2];
    var end = row[3];
    var hours = row[4];
    var shift = row[5];
    var branch = row[6];
    var remark = (row[7] !== undefined && row[7] !== null) ? String(row[7]).trim() : '';
    return (
      '<div class="result-card">' +
        '<div class="result-row"><span class="result-label">姓名</span><span class="result-value">' + escapeHtml(name || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">日期</span><span class="result-value">' + formatDateWithWeekday(date) + '</span></div>' +
        '<div class="result-row"><span class="result-label">班別</span><span class="result-value">' + escapeHtml(shift || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">分店</span><span class="result-value">' + escapeHtml(branch || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">上班</span><span class="result-value">' + escapeHtml(start || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">下班</span><span class="result-value">' + escapeHtml(end || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">時數</span><span class="result-value">' + formatHoursWithMinutes(hours, start, end) + '</span></div>' +
        '<div class="result-row result-row-remark"><span class="result-label">備註</span><textarea class="remark-input" data-type="schedule" data-branch="' + escapeHtmlAttr(branch || '') + '" data-name="' + escapeHtmlAttr(name || '') + '" data-date="' + escapeHtmlAttr(date || '') + '" data-start="' + escapeHtmlAttr(start || '') + '" data-end="' + escapeHtmlAttr(end || '') + '" placeholder="可填寫備註">' + escapeHtml(remark) + '</textarea><button type="button" class="person-btn save-remark-btn">儲存</button></div>' +
      '</div>'
    );
  }).join('');

  scheduleResultSection.querySelectorAll('.save-remark-btn').forEach(function(btn) {
    btn.addEventListener('click', handleSaveRemarkClick);
  });
  scheduleResultSection.classList.add('show');
}

/**
 * 顯示打卡查詢結果
 */
function renderAttendanceResults(result) {
  const details = result.details || {};
  const records = Array.isArray(result.records) ? result.records : [];
  const branchLabel = details.branch ? details.branch : '—';
  const dateRangeLabel = details.startDate
    ? (details.startDate.replace(/-/g, '/') + (details.endDate && details.endDate !== details.startDate ? ' ~ ' + details.endDate.replace(/-/g, '/') : ''))
    : (details.date ? details.date.replace(/-/g, '/') : (details.yearMonth ? details.yearMonth.substring(0,4) + '/' + details.yearMonth.substring(4,6) : '—'));
  scheduleSummary.innerHTML = `
    <div class="summary-item">
      <div class="summary-label">日期範圍</div>
      <div class="summary-value">${dateRangeLabel}</div>
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

  scheduleList.innerHTML = records.map((row, idx) => {
    const branch = row[0];
    const empNo = row[1];
    const empAccount = row[2];
    const name = row[3];
    const date = row[4];
    const start = row[5];
    const end = row[6];
    const hours = row[7];
    const status = row[8];
    const remark = (row[9] !== undefined && row[9] !== null) ? String(row[9]).trim() : '';
    return (
      '<div class="result-card">' +
        '<div class="result-row"><span class="result-label">分店</span><span class="result-value">' + escapeHtml(branch || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">員工編號</span><span class="result-value">' + escapeHtml(empNo || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">員工帳號</span><span class="result-value">' + escapeHtml(empAccount || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">姓名</span><span class="result-value">' + escapeHtml(name || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">打卡日期</span><span class="result-value">' + formatDateWithWeekday(date) + '</span></div>' +
        '<div class="result-row"><span class="result-label">上班</span><span class="result-value">' + escapeHtml(start || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">下班</span><span class="result-value">' + escapeHtml(end || '—') + '</span></div>' +
        '<div class="result-row"><span class="result-label">工作時數</span><span class="result-value">' + formatHoursWithMinutes(hours, start, end) + '</span></div>' +
        (status && String(status).trim() ? '<div class="result-row"><span class="result-label">狀態</span><span class="result-value">' + escapeHtml(status) + '</span></div>' : '') +
        '<div class="result-row result-row-remark"><span class="result-label">備註</span><textarea class="remark-input" data-type="attendance" data-branch="' + escapeHtmlAttr(branch || '') + '" data-emp-account="' + escapeHtmlAttr(empAccount || '') + '" data-date="' + escapeHtmlAttr(date || '') + '" data-start="' + escapeHtmlAttr(start || '') + '" data-end="' + escapeHtmlAttr(end || '') + '" placeholder="可填寫備註">' + escapeHtml(remark) + '</textarea><button type="button" class="person-btn save-remark-btn">儲存</button></div>' +
      '</div>'
    );
  }).join('');

  scheduleResultSection.querySelectorAll('.save-remark-btn').forEach(function(btn) {
    btn.addEventListener('click', handleSaveRemarkClick);
  });
  scheduleResultSection.classList.add('show');
}

/**
 * 載入班表與打卡比對
 */
async function handleLoadCompare() {
  const mode = document.querySelector('input[name="dateFilterMode"]:checked');
  const isMonth = mode && mode.value === 'month';
  let yearMonth = '';
  let startDate = '';
  let endDate = '';
  if (isMonth) {
    var qInp = document.getElementById('queryYearMonthInput');
    var qSel = document.getElementById('queryYearMonthSelect');
    yearMonth = (qInp && qInp.value && qInp.value.trim().match(/^\d{6}$/)) ? qInp.value.trim() : (qSel && qSel.value ? qSel.value : '');
  } else {
    var qStart = document.getElementById('queryStartDate');
    var qEnd = document.getElementById('queryEndDate');
    startDate = qStart && qStart.value ? qStart.value.trim() : '';
    endDate = (qEnd && qEnd.value ? qEnd.value.trim() : '') || startDate;
  }
  if (!yearMonth && (!startDate || startDate.length !== 10)) {
    showAlert('error', '請選擇月份（例如 202601）或日期區間');
    return;
  }
  const queryBranchEl = document.getElementById('queryBranchSelect');
  const branchVal = queryBranchEl && queryBranchEl.value ? queryBranchEl.value.trim() : '';
  if (!branchVal) {
    showAlert('error', '請選擇分店');
    return;
  }
  const names = getSelectedPersonNames();
  const loadCompareBtn = document.getElementById('loadCompareBtn');
  const compareResultSection = document.getElementById('compareResultSection');
  if (loadCompareBtn) loadCompareBtn.disabled = true;
  loadCompareBtn.textContent = '載入中...';
  hideAlert();
  if (compareResultSection) compareResultSection.classList.remove('show');
  showLoadingOverlay();
  var url = CONFIG.GAS_URL + '?action=loadCompare&branch=' + encodeURIComponent(branchVal);
  if (yearMonth) url += '&yearMonth=' + encodeURIComponent(yearMonth);
  if (startDate) url += '&startDate=' + encodeURIComponent(startDate);
  if (endDate) url += '&endDate=' + encodeURIComponent(endDate);
  if (names.length > 0) url += '&names=' + encodeURIComponent(names.join(','));
  try {
    var response = await fetch(url, { method: 'GET', mode: 'cors' });
    var result = await response.json();
    if (!response.ok || !result.success) throw new Error(result.error || '載入失敗');
    renderCompareResults(result.items || []);
    populateComparePersonCheckboxes(result.items || [], {});
    if (compareResultSection) compareResultSection.classList.add('show');
  } catch (error) {
    showAlert('error', '載入比對失敗：' + error.message);
  } finally {
    if (loadCompareBtn) { loadCompareBtn.disabled = false; loadCompareBtn.textContent = '載入比對'; }
    hideLoadingOverlay();
  }
}

/**
 * 從比對結果補足人員複選框（合併比對結果中的人員）
 */
function populateComparePersonCheckboxes(items, existingNames) {
  var group = document.getElementById('personCheckboxGroup');
  if (!group) return;
  existingNames = existingNames || {};
  var nameSet = Object.assign({}, existingNames);
  group.querySelectorAll('input[type="checkbox"]').forEach(function(cb) { if (cb.value) nameSet[cb.value] = true; });
  items.forEach(function(item) {
    var n = item.displayName || (item.attendance && item.attendance.name) || (item.schedule && item.schedule.name);
    if (n) nameSet[n] = true;
  });
  var names = Object.keys(nameSet).sort();
  var checkedNames = [];
  group.querySelectorAll('input[type="checkbox"]:checked').forEach(function(cb) { if (cb.value) checkedNames.push(cb.value); });
  renderQueryPersonCheckboxes(names, { checked: checkedNames });
}

/**
 * 渲染比對區塊人員複選框
 * @param {Array} names - 人員名單
 * @param {Object} opts - { checked: [] } 要預先勾選的名單
 */
function renderComparePersonCheckboxes(names, opts) {
  var comparePersonCheckboxGroup = document.getElementById('comparePersonCheckboxGroup');
  var selectAllComparePersonsBtn = document.getElementById('selectAllComparePersonsBtn');
  var clearAllComparePersonsBtn = document.getElementById('clearAllComparePersonsBtn');
  if (!comparePersonCheckboxGroup) return;
  var checkedSet = {};
  if (opts && Array.isArray(opts.checked)) opts.checked.forEach(function(n) { checkedSet[n] = true; });
  var fromSchedule = opts && opts.fromSchedule;
  comparePersonCheckboxGroup.innerHTML = '';
  if (names.length > 0) {
    if (selectAllComparePersonsBtn) selectAllComparePersonsBtn.disabled = false;
    if (clearAllComparePersonsBtn) clearAllComparePersonsBtn.disabled = false;
    names.forEach(function(n) {
      var label = document.createElement('label');
      var cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = n;
      if (checkedSet[n]) cb.checked = true;
      label.appendChild(cb);
      label.appendChild(document.createTextNode(n));
      comparePersonCheckboxGroup.appendChild(label);
    });
  } else {
    comparePersonCheckboxGroup.innerHTML = '<span class="person-placeholder">' + (fromSchedule ? '此月份打卡無人員' : '此分店無人員資料') + '</span>';
    if (selectAllComparePersonsBtn) selectAllComparePersonsBtn.disabled = true;
    if (clearAllComparePersonsBtn) clearAllComparePersonsBtn.disabled = true;
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
    var s = item.schedule || null;
    var a = item.attendance || null;
    var corr = item.correction || null;
    var displayName = item.displayName || (a && a.name) || (s && s.name) || '—';
    var empAccount = (s && s.empAccount) || (a && a.empAccount) || '';
    var branch = (s && s.branch) || (a && a.branch) || '';
    var date = (s && s.date) || (a && a.date) || '';
    var scheduleStart = s ? (s.startTime || '—') : '—';
    var scheduleEnd = s ? (s.endTime || '—') : '—';
    var scheduleHours = s ? (s.hours || '—') : '—';
    var attendanceStart = a ? (a.startTime || '—') : '—';
    var attendanceEnd = a ? (a.endTime || '—') : '—';
    var attendanceHours = a ? (a.hours || '—') : '—';
    var attendanceStatus = a ? (a.status || '—') : '—';

    var correctedStart = corr ? corr.correctedStart : '';
    var correctedEnd = corr ? corr.correctedEnd : '';
    var correctionRemark = corr ? (corr.remark || '') : '';
    var isCorrected = !!(corr && correctedStart && correctedEnd);
    var scheduleRemark = s ? (s.remark || '') : '';
    var attendanceRemark = a ? (a.remark || '') : '';

    var scheduleText = scheduleStart + '–' + scheduleEnd + ' | ' + formatHoursWithMinutes(scheduleHours, scheduleStart, scheduleEnd);
    var hoursPart = formatHoursWithMinutes(attendanceHours, attendanceStart, attendanceEnd);
    var statusStr = attendanceStatus ? String(attendanceStatus).trim() : '';
    var attendanceText = attendanceStart + '–' + attendanceEnd + ' | ' + hoursPart +
      (statusStr && statusStr !== '—' ? ' | ' + statusStr : '');
    var overtimeAlert = !!(item.overtimeAlert);
    var overlapWarning = !!(item.overlapWarning);
    var confirmedIgnore = !!(item.confirmedIgnore);
    var hasAlert = overtimeAlert || overlapWarning;
    var showConfirmBtn = hasAlert && a && !confirmedIgnore;

    var payload = JSON.stringify({
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
      attendanceStatus: a ? a.status : '',
      scheduleRemark: scheduleRemark,
      attendanceRemark: attendanceRemark,
      correctionRemark: correctionRemark
    });

    return (
      '<div class="compare-card' + (isCorrected ? ' corrected' : '') + (overlapWarning ? ' overlap-warning' : '') + (overtimeAlert ? ' overtime-warning' : '') + '" data-payload="' + escapeHtmlAttr(payload) + '">' +
        (overlapWarning ? '<div class="compare-card-overlap-badge">⚠ 時間重疊</div>' : '') +
        (overtimeAlert ? '<div class="compare-card-overtime-badge">⚠ 加班警示</div>' : '') +
        '<div class="compare-card-header">' +
          escapeHtml(displayName) + '<span class="compare-card-date">' + escapeHtml(formatDateWithWeekday(date)) + '</span>' +
          (confirmedIgnore ? '<span class="compare-card-confirmed-badge">已確認</span><button type="button" class="unconfirm-btn">取消確認</button>' : (showConfirmBtn ? '<button type="button" class="confirm-pending-btn">待確認</button>' : '')) +
          (branch ? '<div class="compare-card-row-label" style="margin-top:4px">' + escapeHtml(branch) + (empAccount ? ' · ' + escapeHtml(empAccount) : '') + '</div>' : '') +
        '</div>' +
        '<div class="compare-card-block">' +
          '<div class="compare-card-block-title">班表</div>' +
          '<div class="compare-card-block-content">' + escapeHtml(scheduleText) + '</div>' +
          (s ? '<div class="compare-card-remark-row"><span class="compare-card-row-label">備註</span><textarea class="schedule-remark-input remark-input" placeholder="可填寫備註" rows="1">' + escapeHtml(scheduleRemark) + '</textarea><button type="button" class="person-btn save-remark-btn" data-type="schedule">儲存</button></div>' : '') +
        '</div>' +
        '<div class="compare-card-block">' +
          '<div class="compare-card-block-title">打卡</div>' +
          '<div class="compare-card-block-content' + (overtimeAlert ? ' overtime-alert' : '') + '">' + escapeHtml(attendanceText) + '</div>' +
          (a ? '<div class="compare-card-remark-row"><span class="compare-card-row-label">備註</span><textarea class="attendance-remark-input remark-input" placeholder="可填寫備註" rows="1">' + escapeHtml(attendanceRemark) + '</textarea><button type="button" class="person-btn save-remark-btn" data-type="attendance">儲存</button></div>' : '') +
        '</div>' +
        '<div class="compare-card-actions">' +
          '<div class="compare-card-actions-row">' +
            '<label><span class="compare-card-row-label">校正上班</span><input type="text" class="corrected-start-input schedule-date-input" placeholder="HH:mm" value="' + escapeHtmlAttr(correctedStart) + '" ' + (isCorrected ? 'readonly' : '') + '></label>' +
            '<label><span class="compare-card-row-label">校正下班</span><input type="text" class="corrected-end-input schedule-date-input" placeholder="HH:mm" value="' + escapeHtmlAttr(correctedEnd) + '" ' + (isCorrected ? 'readonly' : '') + '></label>' +
          '</div>' +
          '<div class="compare-card-remark-row"><span class="compare-card-row-label">校正備註</span><textarea class="correction-remark-input remark-input" placeholder="可填寫備註" rows="1">' + escapeHtml(correctionRemark) + '</textarea></div>' +
          (isCorrected
            ? '<div style="display:flex;gap:10px;align-items:center"><span class="compare-card-badge">已校正</span><button type="button" class="person-btn edit-correction-btn">編輯</button></div>'
            : '<button type="button" class="load-schedule-btn submit-correction-btn">送出校正</button>') +
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
  compareList.querySelectorAll('.save-remark-btn').forEach(function(btn) {
    btn.addEventListener('click', handleSaveRemarkClick);
  });
  compareList.querySelectorAll('.confirm-pending-btn').forEach(function(btn) {
    btn.addEventListener('click', handleConfirmIgnoreClick);
  });
  compareList.querySelectorAll('.unconfirm-btn').forEach(function(btn) {
    btn.addEventListener('click', handleUnconfirmIgnoreClick);
  });
}

/**
 * 處理打卡警示確認按鈕點擊（待確認）
 */
function handleConfirmIgnoreClick(e) {
  var btn = e.target;
  var card = btn.closest('.compare-card');
  if (!card) return;
  var payloadStr = card.getAttribute('data-payload');
  if (!payloadStr) return;
  showConfirm('確定要將此筆打卡警示標記為已確認？', function() {
  try {
    var payload = JSON.parse(payloadStr);
    doConfirmIgnoreAttendance({
      branch: payload.branch,
      empAccount: payload.empAccount,
      date: payload.date,
      attendanceStart: payload.attendanceStart,
      attendanceEnd: payload.attendanceEnd
    });
  } catch (err) {
    showAlert('error', '資料格式錯誤');
  }
  });
}

/**
 * 處理取消確認按鈕點擊
 */
function handleUnconfirmIgnoreClick(e) {
  var btn = e.target;
  var card = btn.closest('.compare-card');
  if (!card) return;
  var payloadStr = card.getAttribute('data-payload');
  if (!payloadStr) return;
  showConfirm('確定要取消確認？將還原為待確認狀態。', function() {
  try {
    var payload = JSON.parse(payloadStr);
    doUnconfirmIgnoreAttendance({
      branch: payload.branch,
      empAccount: payload.empAccount,
      date: payload.date,
      attendanceStart: payload.attendanceStart,
      attendanceEnd: payload.attendanceEnd
    });
  } catch (err) {
    showAlert('error', '資料格式錯誤');
  }
  });
}

/**
 * 將 HHMM 轉為 HH:mm（相容 0530、05:30）
 */
function normalizeTimeInput(val) {
  if (!val || typeof val !== 'string') return val;
  var s = val.trim();
  if (/^\d{4}$/.test(s)) return s.substring(0, 2) + ':' + s.substring(2);
  return s;
}

/**
 * 日期加上星期幾，格式：YYYY/MM/DD (一) ~ (日)
 * @param {string} dateStr - YYYY-MM-DD 或 YYYY/MM/DD
 */
function formatDateWithWeekday(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return '—';
  var s = dateStr.trim();
  if (!s) return '—';
  var parts = s.split(/[-/]/);
  if (parts.length < 3) return s;
  var y = parseInt(parts[0], 10);
  var m = parseInt(parts[1], 10) - 1;
  var d = parseInt(parts[2], 10);
  if (isNaN(y) || isNaN(m) || isNaN(d)) return s;
  try {
    var dt = new Date(y, m, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== m || dt.getDate() !== d) return s;
    var weekdays = ['日', '一', '二', '三', '四', '五', '六'];
    var wd = weekdays[dt.getDay()];
    var display = y + '/' + String(m + 1).padStart(2, '0') + '/' + String(d).padStart(2, '0');
    return display + ' (' + wd + ')';
  } catch (e) {
    return s;
  }
}

/**
 * 將時間字串（HH:mm 或 HHmm）或 Date 轉為當日分鐘數（0–1439），失敗回傳 null
 */
function parseTimeToMinutes(val) {
  if (val === undefined || val === null) return null;
  if (val instanceof Date) {
    var h = val.getHours();
    var m = val.getMinutes();
    return h * 60 + m;
  }
  var s = String(val).trim();
  if (!s) return null;
  var parts = s.match(/^(\d{1,2}):(\d{2})$/) || (s.match(/^(\d{4})$/) ? [null, s.substring(0, 2), s.substring(2, 4)] : null);
  if (!parts) return null;
  var h = parseInt(parts[1], 10);
  var m = parseInt(parts[2], 10);
  if (isNaN(h) || isNaN(m) || h < 0 || h > 23 || m < 0 || m > 59) return null;
  return h * 60 + m;
}

/**
 * 從 start、end 計算實際分鐘數，跨日處理。失敗回傳 null。
 */
function calcMinutesFromTimeRange(startStr, endStr) {
  var startMins = parseTimeToMinutes(startStr);
  var endMins = parseTimeToMinutes(endStr);
  if (startMins === null || endMins === null) return null;
  var diff = endMins - startMins;
  if (diff < 0) diff += 24 * 60;
  return diff;
}

/**
 * 格式化工作時數為「n.m 小時 (iii 分)」，分鐘依 start/end 實際計算（方案 B）。
 * 無 start/end 時退回 hours * 60。
 */
function formatHoursWithMinutes(hours, start, end) {
  var actualMinutes = (start && end) ? calcMinutesFromTimeRange(start, end) : null;
  if (actualMinutes === null) {
    if (hours === undefined || hours === null || String(hours).trim() === '') return '—';
    var s = String(hours).trim();
    var m = s.match(/[\d.]+/);
    if (!m) return s;
    var num = parseFloat(m[0]);
    if (isNaN(num)) return s;
    actualMinutes = Math.round(num * 60);
  }
  var hoursVal = actualMinutes / 60;
  var hoursStr = hoursVal % 1 === 0 ? String(Math.round(hoursVal)) : (Math.round(hoursVal * 10) / 10).toString();
  return hoursStr + '小時(' + actualMinutes + '分)';
}

function escapeHtml(s) {
  if (s === undefined || s === null) return '';
  var t = String(s);
  return t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function escapeHtmlAttr(s) {
  if (s === undefined || s === null) return '';
  return escapeHtml(s).replace(/'/g, '&#39;');
}

/**
 * 處理備註儲存按鈕點擊
 */
function handleSaveRemarkClick(e) {
  var btn = e.target;
  var row = btn.closest('.result-row-remark') || btn.closest('.compare-card-remark-row');
  var textarea = row ? row.querySelector('.remark-input') : null;
  if (!textarea) return;
  var type = textarea.getAttribute('data-type') || btn.getAttribute('data-type');
  var remark = (textarea.value || '').trim();
  var payload = { remark: remark };
  if (type === 'schedule') {
    if (textarea.getAttribute('data-branch') !== null) {
      payload.branch = textarea.getAttribute('data-branch') || '';
      payload.name = textarea.getAttribute('data-name') || '';
      payload.date = textarea.getAttribute('data-date') || '';
      payload.start = textarea.getAttribute('data-start') || '';
      payload.end = textarea.getAttribute('data-end') || '';
    } else {
      var card = btn.closest('.compare-card');
      if (!card || !card.getAttribute('data-payload')) return;
      try {
        var p = JSON.parse(card.getAttribute('data-payload'));
        payload.branch = p.branch || '';
        payload.name = p.displayName || p.name || '';
        payload.date = p.date || '';
        payload.start = p.scheduleStart || '';
        payload.end = p.scheduleEnd || '';
      } catch (err) { return; }
    }
    doUpdateScheduleRemark(payload);
  } else if (type === 'attendance') {
    if (textarea.getAttribute('data-branch') !== null) {
      payload.branch = textarea.getAttribute('data-branch') || '';
      payload.empAccount = textarea.getAttribute('data-emp-account') || '';
      payload.date = textarea.getAttribute('data-date') || '';
      payload.start = textarea.getAttribute('data-start') || '';
      payload.end = textarea.getAttribute('data-end') || '';
    } else {
      var card = btn.closest('.compare-card');
      if (!card || !card.getAttribute('data-payload')) return;
      try {
        var p = JSON.parse(card.getAttribute('data-payload'));
        payload.branch = p.branch || '';
        payload.empAccount = p.empAccount || '';
        payload.date = p.date || '';
        payload.start = p.attendanceStart || '';
        payload.end = p.attendanceEnd || '';
      } catch (err) { return; }
    }
    doUpdateAttendanceRemark(payload);
  }
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
  const correctionRemarkInput = card.querySelector('.correction-remark-input');
  var correctedStart = correctedStartInput ? correctedStartInput.value.trim() : '';
  var correctedEnd = correctedEndInput ? correctedEndInput.value.trim() : '';
  var correctionRemark = correctionRemarkInput ? correctionRemarkInput.value.trim() : '';
  if (!correctedStart || !correctedEnd) {
    showAlert('error', '請填寫校正上班時間與校正下班時間');
    return;
  }
  payload.correctedStart = normalizeTimeInput(correctedStart);
  payload.correctedEnd = normalizeTimeInput(correctedEnd);
  payload.remark = correctionRemark;
  payload.correctionRemark = correctionRemark;
  doSubmitCorrection(payload);
}

/**
 * 處理編輯按鈕點擊（已校正狀態下切換為可編輯）
 */
function handleEditCorrectionClick(e) {
  var btn = e.target;
  var card = btn.closest('.compare-card');
  if (!card) return;
  var correctedStartInput = card.querySelector('.corrected-start-input');
  var correctedEndInput = card.querySelector('.corrected-end-input');
  if (correctedStartInput) correctedStartInput.removeAttribute('readonly');
  if (correctedEndInput) correctedEndInput.removeAttribute('readonly');
  var btnRow = btn.closest('div');
  var newBtn = document.createElement('button');
  newBtn.type = 'button';
  newBtn.className = 'load-schedule-btn submit-correction-btn';
  newBtn.textContent = '送出校正';
  newBtn.addEventListener('click', handleSubmitCorrectionClick);
  if (btnRow && btnRow.parentNode) {
    btnRow.replaceWith(newBtn);
  } else {
    btn.replaceWith(newBtn);
  }
}

/**
 * 送出校正到 API
 */
async function doSubmitCorrection(payload) {
  try {
    showLoadingOverlay();
    const response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
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
        correctedEnd: payload.correctedEnd,
        remark: payload.remark || payload.correctionRemark || '',
        correctionRemark: payload.remark || payload.correctionRemark || ''
      })
    });
    var result;
    var text = await response.text();
    try {
      result = text ? JSON.parse(text) : {};
    } catch (parseErr) {
      if (!response.ok) {
        showAlert('error', '校正送出失敗：伺服器回傳錯誤 (HTTP ' + response.status + ')');
      } else {
        showAlert('error', '校正送出失敗：伺服器未回傳有效資料');
      }
      return;
    }
    if (!response.ok || !result.success) {
      showAlert('error', '校正送出失敗：' + (result.error || '未知錯誤'));
      return;
    }
    showAlert('success', '校正紀錄已送出');
    const loadCompareBtn = document.getElementById('loadCompareBtn');
    if (loadCompareBtn) loadCompareBtn.click();
  } catch (error) {
    showAlert('error', '校正送出失敗：' + (error.message || '網路或連線錯誤'));
  } finally {
    hideLoadingOverlay();
  }
}

/**
 * 更新班表備註到 API
 */
async function doUpdateScheduleRemark(payload) {
  try {
    showLoadingOverlay();
    var response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
      body: JSON.stringify({
        action: 'updateScheduleRemark',
        branch: payload.branch,
        name: payload.name,
        date: payload.date,
        start: payload.start,
        end: payload.end,
        remark: payload.remark
      })
    });
    var result = await response.json();
    if (!response.ok || !result.success) {
      throw new Error(result.error || '更新備註失敗');
    }
    showAlert('success', '備註已更新');
  } catch (error) {
    showAlert('error', '更新班表備註失敗：' + error.message);
  } finally {
    hideLoadingOverlay();
  }
}

/**
 * 更新打卡備註到 API
 */
async function doUpdateAttendanceRemark(payload) {
  try {
    showLoadingOverlay();
    var response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
      body: JSON.stringify({
        action: 'updateAttendanceRemark',
        branch: payload.branch,
        empAccount: payload.empAccount,
        date: payload.date,
        start: payload.start,
        end: payload.end,
        remark: payload.remark
      })
    });
    var result = await response.json();
    if (!response.ok || !result.success) {
      throw new Error(result.error || '更新備註失敗');
    }
    showAlert('success', '備註已更新');
  } catch (error) {
    showAlert('error', '更新打卡備註失敗：' + error.message);
  } finally {
    hideLoadingOverlay();
  }
}

/**
 * 打卡警示確認到 API
 */
async function doConfirmIgnoreAttendance(payload) {
  try {
    showLoadingOverlay();
    var response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
      body: JSON.stringify({
        action: 'confirmIgnoreAttendance',
        branch: payload.branch,
        empAccount: payload.empAccount,
        date: payload.date,
        attendanceStart: payload.attendanceStart,
        attendanceEnd: payload.attendanceEnd
      })
    });
    var result = await response.json();
    if (!response.ok || !result.success) {
      throw new Error(result.error || '確認失敗');
    }
    showAlert('success', '已確認');
    var loadCompareBtn = document.getElementById('loadCompareBtn');
    if (loadCompareBtn) loadCompareBtn.click();
  } catch (error) {
    showAlert('error', '確認失敗：' + error.message);
  } finally {
    hideLoadingOverlay();
  }
}

/**
 * 取消打卡警示確認到 API
 */
async function doUnconfirmIgnoreAttendance(payload) {
  try {
    showLoadingOverlay();
    var response = await fetch(CONFIG.GAS_URL, {
      method: 'POST',
      mode: 'cors',
      headers: { 'Content-Type': 'text/plain; charset=utf-8' },
      body: JSON.stringify({
        action: 'unconfirmIgnoreAttendance',
        branch: payload.branch,
        empAccount: payload.empAccount,
        date: payload.date,
        attendanceStart: payload.attendanceStart,
        attendanceEnd: payload.attendanceEnd
      })
    });
    var result = await response.json();
    if (!response.ok || !result.success) {
      throw new Error(result.error || '取消確認失敗');
    }
    showAlert('success', '已取消確認');
    var loadCompareBtn = document.getElementById('loadCompareBtn');
    if (loadCompareBtn) loadCompareBtn.click();
  } catch (error) {
    showAlert('error', '取消確認失敗：' + error.message);
  } finally {
    hideLoadingOverlay();
  }
}
