/**
 * 薪資計算系統 - 環境設定
 * 測試環境
 */

/**
 * 打卡 sheet 欄位 index（0-based），共 14 欄
 * A-K: 分店,員工編號,員工帳號,員工姓名,打卡日期,上班,下班,時數,狀態,備註,是否有效,校正備註,建立時間,校正時間
 */
var ATTENDANCE_COL = {
  BRANCH: 0, EMP_NO: 1, EMP_ACCOUNT: 2, NAME: 3, DATE: 4,
  START: 5, END: 6, HOURS: 7, STATUS: 8, REMARK: 9,
  VALID: 10, CORRECTION_REMARK: 11, CREATED_AT: 12, CORRECTED_AT: 13
};

/**
 * 班表 sheet 欄位 index（0-based），共 10 欄
 * A-J: 員工姓名,排班日期,上班,下班,時數,班別,分店,備註,建立時間,修改時間
 */
var SCHEDULE_COL = {
  NAME: 0, DATE: 1, START: 2, END: 3, HOURS: 4, SHIFT: 5,
  BRANCH: 6, REMARK: 7, CREATED_AT: 8, MODIFIED_AT: 9
};

// 從 Script Properties 讀取設定
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  
  return {
    // Google Sheets ID（從 Script Properties 讀取）
    SHEET_ID: props.getProperty('SHEET_ID') || '1YiwEZdZGGt6XISbKUexnQK-uaEMwy1TN1frsbyvc0f8',
    
    // 環境標識
    ENVIRONMENT: 'test',
    
    // 工作表名稱對應
    SHEET_NAMES: {
      SCHEDULE: '班表',          // 班表資料
      BRANCH: '分店',            // 分店清單（代碼、名稱、啟用狀態、排序、打卡地點）
      ATTENDANCE: '打卡',        // 打卡紀錄
      PERSONNEL: '人員',         // 人員（員工帳號、班表名稱、打卡名稱、分店）
      CORRECTION: '校正',        // 班表與打卡校正紀錄
      CALCULATION: '計算結果',   // 計算結果
      ADJUSTMENTS: '調整記錄',   // 調整記錄
      LOGS: '處理記錄'          // 處理記錄
    },
    
    // 檔案大小限制（bytes）
    MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
    
    // 支援的 Excel 工作表名稱（用於自動偵測）
    SUPPORTED_SHEET_NAMES: ['11501', '班表', 'Schedule', 'Attendance'],
    
    // 加班警示閾值（分鐘）。打卡時數 − 班表時數 > 此值時，前端將打卡區塊標紅。從 Script Property OVERTIME_ALERT 讀取，未設定時 0＝不警示。
    OVERTIME_ALERT: (function() {
      var v = props.getProperty('OVERTIME_ALERT');
      if (v === null || v === undefined || String(v).trim() === '') return 0;
      var n = parseInt(String(v).trim(), 10);
      return isNaN(n) ? 0 : n;
    })()
  };
}

/**
 * 取得 Google Sheets 物件
 */
function getSpreadsheet() {
  const config = getConfig();
  try {
    return SpreadsheetApp.openById(config.SHEET_ID);
  } catch (error) {
    Logger.log('無法開啟 Google Sheets: ' + error.message);
    throw new Error('Google Sheets ID 設定錯誤，請檢查 Script Properties 中的 SHEET_ID');
  }
}

/**
 * 取得工作表（若不存在則回傳 null，不建立新工作表）
 */
function getSheet(sheetName) {
  const ss = getSpreadsheet();
  return ss.getSheetByName(sheetName);
}

/**
 * 取得或建立指定名稱的工作表（僅用於寫入情境）
 */
function getOrCreateSheet(sheetName) {
  const sheet = getSheet(sheetName);
  if (sheet) return sheet;
  const ss = getSpreadsheet();
  Logger.log('建立新工作表: ' + sheetName);
  return ss.insertSheet(sheetName);
}

/**
 * Log 等級定義
 */
const LOG_LEVEL = {
  OFF: 0,      // 關閉 Log
  OPERATION: 1, // 營運等級（重要操作、錯誤）
  DEBUG: 2     // Debug 等級（所有操作、詳細資訊）
};

/**
 * 取得當前 Log 等級設定
 * @return {number} Log 等級（0=關閉, 1=營運, 2=Debug）
 */
function getLogLevel() {
  const props = PropertiesService.getScriptProperties();
  const logLevel = props.getProperty('Log_Level');
  
  if (logLevel === null || logLevel === undefined || logLevel === '') {
    return LOG_LEVEL.DEBUG; // 預設為 DEBUG 等級
  }
  
  return parseInt(logLevel, 10);
}

/**
 * 記錄日誌到 Google Sheets
 * @param {string} message - 日誌訊息
 * @param {string} level - 日誌等級 ('DEBUG', 'INFO', 'WARNING', 'ERROR', 'OPERATION')
 * @param {Object} details - 額外的詳細資訊（選填）
 */
function logToSheet(message, level = 'INFO', details = null) {
  try {
    // 取得當前 Log 等級設定
    const currentLogLevel = getLogLevel();
    
    // 如果 Log 關閉，直接返回
    if (currentLogLevel === LOG_LEVEL.OFF) {
      return;
    }
    
    // 判斷是否應該記錄此等級的 Log
    const shouldLog = checkShouldLog(level, currentLogLevel);
    if (!shouldLog) {
      return;
    }
    
    const config = getConfig();
    const sheet = getOrCreateSheet('Log'); // 固定使用 'Log' 作為 sheet 名稱
    
    // 如果是第一次寫入，建立標題列
    if (sheet.getLastRow() === 0) {
      initLogSheet(sheet);
    }
    
    const timestamp = new Date();
    const logEntry = [
      timestamp,                          // A: 時間
      level,                              // B: 等級
      message,                            // C: 訊息
      config.ENVIRONMENT,                 // D: 環境
      'system',                           // E: 使用者（改為固定值，避免權限問題）
      details ? JSON.stringify(details) : '' // F: 詳細資訊
    ];
    
    sheet.appendRow(logEntry);
    
    // 同時輸出到 Apps Script Logger（方便開發時查看）
    Logger.log(`[${level}] ${message}`);
    if (details) {
      Logger.log('Details: ' + JSON.stringify(details));
    }
    
  } catch (error) {
    // 如果記錄日誌失敗，至少輸出到 Logger
    Logger.log('記錄日誌失敗: ' + error.message);
    Logger.log('原始訊息: ' + message);
  }
}

/**
 * 判斷是否應該記錄此等級的 Log
 * @param {string} level - 日誌等級
 * @param {number} currentLogLevel - 當前設定的 Log 等級
 * @return {boolean} 是否應該記錄
 */
function checkShouldLog(level, currentLogLevel) {
  // 營運等級 (1)：只記錄 OPERATION 和 ERROR
  if (currentLogLevel === LOG_LEVEL.OPERATION) {
    return (level === 'OPERATION' || level === 'ERROR');
  }
  
  // Debug 等級 (2)：記錄所有
  if (currentLogLevel === LOG_LEVEL.DEBUG) {
    return true;
  }
  
  // 其他情況不記錄
  return false;
}

/**
 * 初始化 Log Sheet（建立標題列並格式化）
 * @param {Sheet} sheet - Log 工作表
 */
function initLogSheet(sheet) {
  // 設定標題列
  const headers = ['時間', '等級', '訊息', '環境', '使用者', '詳細資訊'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // 格式化標題列
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#434343');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  headerRange.setHorizontalAlignment('center');
  
  // 凍結標題列
  sheet.setFrozenRows(1);
  
  // 設定欄寬
  sheet.setColumnWidth(1, 150); // 時間
  sheet.setColumnWidth(2, 100); // 等級
  sheet.setColumnWidth(3, 300); // 訊息
  sheet.setColumnWidth(4, 80);  // 環境
  sheet.setColumnWidth(5, 150); // 使用者
  sheet.setColumnWidth(6, 400); // 詳細資訊
  
  Logger.log('Log Sheet 初始化完成');
}

/**
 * 快速記錄各等級的 Log（便利函數）
 */
function logDebug(message, details = null) {
  logToSheet(message, 'DEBUG', details);
}

function logInfo(message, details = null) {
  logToSheet(message, 'INFO', details);
}

function logWarning(message, details = null) {
  logToSheet(message, 'WARNING', details);
}

function logError(message, details = null) {
  logToSheet(message, 'ERROR', details);
}

function logOperation(message, details = null) {
  logToSheet(message, 'OPERATION', details);
}

/**
 * 清理舊的 Log（保留最近 N 天）
 * @param {number} daysToKeep - 保留天數（預設 30 天）
 */
function cleanOldLogs(daysToKeep = 30) {
  try {
    const sheet = getOrCreateSheet('Log');
    const lastRow = sheet.getLastRow();
    
    if (lastRow <= 1) {
      logInfo('沒有需要清理的 Log');
      return;
    }
    
    const data = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // 讀取時間欄
    const cutoffDate = new Date();
    cutoffDate.setDate(cutoffDate.getDate() - daysToKeep);
    
    let deleteCount = 0;
    for (let i = data.length - 1; i >= 0; i--) {
      const logDate = new Date(data[i][0]);
      if (logDate < cutoffDate) {
        sheet.deleteRow(i + 2); // +2 因為有標題列且陣列從 0 開始
        deleteCount++;
      }
    }
    
    logOperation(`清理 Log 完成，刪除 ${deleteCount} 筆舊資料（${daysToKeep} 天前）`);
    
  } catch (error) {
    logError('清理舊 Log 失敗: ' + error.message);
  }
}
