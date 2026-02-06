/**
 * 測試工具函數
 * 這些函數只用於測試與除錯，不會在正式流程中呼叫
 */

/**
 * 匯出「國安班表」的結構與資料
 * 在 GAS 編輯器中執行此函數，然後查看執行記錄
 */
function exportSheetStructure() {
  try {
    const ss = SpreadsheetApp.openById('1YiwEZdZGGt6XISbKUexnQK-uaEMwy1TN1frsbyvc0f8');
    const sheet = ss.getSheetByName('國安班表');
    
    if (!sheet) {
      Logger.log('錯誤：找不到「國安班表」工作表');
      Logger.log('可用的工作表：');
      ss.getSheets().forEach(s => Logger.log('  - ' + s.getName()));
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    Logger.log('=== 國安班表結構 ===');
    Logger.log('最後一列：' + lastRow);
    Logger.log('最後一欄：' + lastCol);
    Logger.log('');
    
    if (lastRow === 0 || lastCol === 0) {
      Logger.log('工作表是空的');
      return;
    }
    
    // 讀取前 10 列資料（含標題）
    const rowsToRead = Math.min(lastRow, 10);
    const data = sheet.getRange(1, 1, rowsToRead, lastCol).getValues();
    
    Logger.log('=== 標題列（第 1 列）===');
    Logger.log(JSON.stringify(data[0]));
    Logger.log('');
    
    Logger.log('=== 欄位對應 ===');
    for (let i = 0; i < data[0].length; i++) {
      const colLetter = columnToLetter(i + 1);
      Logger.log(colLetter + ' 欄 = ' + data[0][i]);
    }
    Logger.log('');
    
    Logger.log('=== 前幾列資料範例 ===');
    for (let i = 1; i < data.length; i++) {
      Logger.log('第 ' + (i + 1) + ' 列: ' + JSON.stringify(data[i]));
    }
    Logger.log('');
    
    Logger.log('=== 完整資料（JSON 格式）===');
    Logger.log(JSON.stringify(data, null, 2));
    
  } catch (error) {
    Logger.log('錯誤：' + error.message);
    Logger.log('請確認：');
    Logger.log('1. Google Sheets ID 是否正確');
    Logger.log('2. GAS 是否有權限存取該 Sheets');
  }
}

/**
 * 列出所有工作表名稱
 */
function listAllSheets() {
  try {
    const ss = SpreadsheetApp.openById('1YiwEZdZGGt6XISbKUexnQK-uaEMwy1TN1frsbyvc0f8');
    
    Logger.log('=== Google Sheets 資訊 ===');
    Logger.log('試算表名稱：' + ss.getName());
    Logger.log('試算表 ID：' + ss.getId());
    Logger.log('');
    
    Logger.log('=== 所有工作表 ===');
    const sheets = ss.getSheets();
    sheets.forEach((sheet, index) => {
      Logger.log((index + 1) + '. ' + sheet.getName() + 
        ' (' + sheet.getLastRow() + ' 列 x ' + 
        sheet.getLastColumn() + ' 欄)');
    });
    
  } catch (error) {
    Logger.log('錯誤：' + error.message);
  }
}

/**
 * 測試讀取上傳的 Excel 原始資料
 */
function testReadUploadedExcel() {
  try {
    const ss = SpreadsheetApp.openById('1YiwEZdZGGt6XISbKUexnQK-uaEMwy1TN1frsbyvc0f8');
    
    // 假設上傳的原始資料會先存在某個工作表
    // 這裡列出所有可能的資料來源
    Logger.log('=== 搜尋可能的資料來源 ===');
    
    const sheets = ss.getSheets();
    sheets.forEach(sheet => {
      const name = sheet.getName();
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      if (lastRow > 0) {
        Logger.log('');
        Logger.log('工作表：' + name);
        Logger.log('大小：' + lastRow + ' 列 x ' + lastCol + ' 欄');
        
        // 讀取前 3 列
        const data = sheet.getRange(1, 1, Math.min(3, lastRow), lastCol).getValues();
        Logger.log('前 3 列資料：');
        data.forEach((row, i) => {
          Logger.log('  第 ' + (i + 1) + ' 列: ' + JSON.stringify(row));
        });
      }
    });
    
  } catch (error) {
    Logger.log('錯誤：' + error.message);
  }
}

/**
 * 測試基本連線與設定
 */
function testBasicSetup() {
  Logger.log('=== 測試基本設定 ===');
  Logger.log('');
  
  // 1. 測試 Config
  Logger.log('1. 測試 Config');
  try {
    const config = getConfig();
    Logger.log('✓ Config 讀取成功');
    Logger.log('  SHEET_ID: ' + config.SHEET_ID);
    Logger.log('  ENVIRONMENT: ' + config.ENVIRONMENT);
    Logger.log('');
  } catch (error) {
    Logger.log('✗ Config 讀取失敗: ' + error.message);
    Logger.log('');
  }
  
  // 2. 測試 Google Sheets 連線
  Logger.log('2. 測試 Google Sheets 連線');
  try {
    const ss = getSpreadsheet();
    Logger.log('✓ Google Sheets 連線成功');
    Logger.log('  試算表名稱: ' + ss.getName());
    Logger.log('  試算表 ID: ' + ss.getId());
    Logger.log('');
  } catch (error) {
    Logger.log('✗ Google Sheets 連線失敗: ' + error.message);
    Logger.log('');
  }
  
  // 3. 測試 Log Level
  Logger.log('3. 測試 Log Level');
  try {
    const logLevel = getLogLevel();
    Logger.log('✓ Log Level: ' + logLevel);
    Logger.log('  0=關閉, 1=營運, 2=Debug');
    Logger.log('');
  } catch (error) {
    Logger.log('✗ Log Level 讀取失敗: ' + error.message);
    Logger.log('');
  }
  
  // 4. 測試寫入 Log
  Logger.log('4. 測試寫入 Log');
  try {
    logOperation('測試記錄 - ' + new Date().toLocaleString('zh-TW'), {
      test: true,
      timestamp: new Date().toISOString()
    });
    Logger.log('✓ Log 寫入成功（請檢查 Google Sheets 的 Log 工作表）');
    Logger.log('');
  } catch (error) {
    Logger.log('✗ Log 寫入失敗: ' + error.message);
    Logger.log('');
  }
  
  Logger.log('=== 測試完成 ===');
  Logger.log('');
  Logger.log('下一步：');
  Logger.log('1. 執行 listAllSheets() 查看所有工作表');
  Logger.log('2. 執行 exportSheetStructure() 查看「國安班表」結構');
}
