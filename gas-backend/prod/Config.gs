/**
 * 薪資計算系統 - 環境設定
 * 測試環境
 */

// 從 Script Properties 讀取設定
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  
  return {
    // Google Sheets ID（從 Script Properties 讀取）
    SHEET_ID: props.getProperty('SHEET_ID') || 'YOUR_TEST_SHEET_ID',
    
    // 環境標識
    ENVIRONMENT: 'test',
    
    // 工作表名稱對應
    SHEET_NAMES: {
      SCHEDULE: '國安班表',      // 班表資料
      ATTENDANCE: '打卡紀錄',    // 打卡紀錄
      CALCULATION: '計算結果',   // 計算結果
      ADJUSTMENTS: '調整記錄',   // 調整記錄
      LOGS: '處理記錄'          // 處理記錄
    },
    
    // 檔案大小限制（bytes）
    MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
    
    // 支援的 Excel 工作表名稱（用於自動偵測）
    SUPPORTED_SHEET_NAMES: ['11501', '班表', 'Schedule', 'Attendance']
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
 * 取得或建立指定名稱的工作表
 */
function getOrCreateSheet(sheetName) {
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    Logger.log('建立新工作表: ' + sheetName);
    sheet = ss.insertSheet(sheetName);
  }
  
  return sheet;
}

/**
 * 記錄處理日誌
 */
function logToSheet(message, level = 'INFO') {
  try {
    const config = getConfig();
    const sheet = getOrCreateSheet(config.SHEET_NAMES.LOGS);
    
    const timestamp = new Date();
    const logEntry = [
      timestamp,
      level,
      message,
      config.ENVIRONMENT
    ];
    
    sheet.appendRow(logEntry);
    
    // 如果是第一列，加入標題
    if (sheet.getLastRow() === 1) {
      sheet.getRange(1, 1, 1, 4).setValues([['時間', '等級', '訊息', '環境']]);
      sheet.getRange(2, 1, 1, 4).setValues([logEntry]);
    }
  } catch (error) {
    Logger.log('記錄日誌失敗: ' + error.message);
  }
}
