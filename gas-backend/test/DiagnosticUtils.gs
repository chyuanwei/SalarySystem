/**
 * 診斷工具 - 用於排查資料寫入問題
 */

/**
 * 診斷完整的上傳流程
 */
function diagnoseUploadFlow() {
  Logger.log('========== 診斷開始 ==========');
  
  // 1. 檢查環境設定
  Logger.log('\n【步驟 1】檢查環境設定');
  const sheetId = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  const logLevel = PropertiesService.getScriptProperties().getProperty('Log_Level');
  
  Logger.log(`SHEET_ID: ${sheetId ? '✓ 已設定' : '✗ 未設定'}`);
  Logger.log(`Log_Level: ${logLevel || '未設定（預設 DEBUG）'}`);
  
  if (!sheetId) {
    Logger.log('錯誤：SHEET_ID 未設定！請先設定 Script Properties');
    return;
  }
  
  // 2. 檢查 Google Sheets 存取
  Logger.log('\n【步驟 2】檢查 Google Sheets 存取');
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    Logger.log(`✓ 成功開啟試算表: ${ss.getName()}`);
    Logger.log(`試算表 ID: ${ss.getId()}`);
    Logger.log(`試算表網址: ${ss.getUrl()}`);
    
    const sheets = ss.getSheets();
    Logger.log(`現有工作表數量: ${sheets.length}`);
    sheets.forEach((sheet, idx) => {
      Logger.log(`  ${idx + 1}. ${sheet.getName()} (${sheet.getLastRow()} 列)`);
    });
  } catch (error) {
    Logger.log(`✗ 無法開啟試算表: ${error.message}`);
    return;
  }
  
  // 3. 測試建立/取得工作表
  Logger.log('\n【步驟 3】測試建立/取得工作表');
  const testSheetName = '診斷測試_' + new Date().getTime();
  try {
    const sheet = getOrCreateSheet(testSheetName);
    Logger.log(`✓ 成功建立工作表: ${sheet.getName()}`);
    
    // 4. 測試寫入資料
    Logger.log('\n【步驟 4】測試寫入資料');
    const testData = [
      ['員工姓名', '排班日期', '上班時間', '下班時間', '工作時數'],
      ['測試員工A', '2026/01/01', '09:00', '18:00', '8.0'],
      ['測試員工B', '2026/01/02', '10:00', '19:00', '9.0']
    ];
    
    const writeResult = appendToSheet(testData, testSheetName);
    Logger.log(`寫入結果: ${JSON.stringify(writeResult)}`);
    
    if (writeResult.success) {
      Logger.log(`✓ 成功寫入 ${writeResult.rowCount} 列資料`);
      
      // 5. 驗證寫入結果
      Logger.log('\n【步驟 5】驗證寫入結果');
      const readData = readFromSheet(testSheetName);
      Logger.log(`讀取到 ${readData.length} 列資料`);
      readData.forEach((row, idx) => {
        Logger.log(`  第 ${idx + 1} 列: ${JSON.stringify(row)}`);
      });
    } else {
      Logger.log(`✗ 寫入失敗: ${writeResult.error}`);
    }
    
    // 6. 清理測試工作表
    Logger.log('\n【步驟 6】清理測試工作表');
    const ss = getSpreadsheet();
    ss.deleteSheet(sheet);
    Logger.log(`✓ 已刪除測試工作表: ${testSheetName}`);
    
  } catch (error) {
    Logger.log(`✗ 測試過程發生錯誤: ${error.message}`);
    Logger.log(`錯誤堆疊: ${error.stack}`);
  }
  
  Logger.log('\n========== 診斷結束 ==========');
}

/**
 * 檢查特定工作表的資料
 */
function checkSheetData(sheetName) {
  if (!sheetName) {
    sheetName = '班表'; // 預設檢查班表
  }
  
  Logger.log(`========== 檢查工作表: ${sheetName} ==========`);
  
  try {
    const ss = getSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      Logger.log(`✗ 工作表不存在: ${sheetName}`);
      Logger.log('\n現有工作表:');
      ss.getSheets().forEach((s, idx) => {
        Logger.log(`  ${idx + 1}. ${s.getName()}`);
      });
      return;
    }
    
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    Logger.log(`✓ 工作表存在`);
    Logger.log(`最後一列: ${lastRow}`);
    Logger.log(`最後一欄: ${lastCol}`);
    
    if (lastRow === 0) {
      Logger.log('⚠ 工作表是空的，沒有任何資料');
      return;
    }
    
    // 顯示前 10 列資料
    const displayRows = Math.min(lastRow, 10);
    Logger.log(`\n前 ${displayRows} 列資料:`);
    
    const data = sheet.getRange(1, 1, displayRows, lastCol).getValues();
    data.forEach((row, idx) => {
      Logger.log(`第 ${idx + 1} 列: ${JSON.stringify(row)}`);
    });
    
    if (lastRow > 10) {
      Logger.log(`... 還有 ${lastRow - 10} 列資料未顯示`);
    }
    
  } catch (error) {
    Logger.log(`✗ 檢查失敗: ${error.message}`);
    Logger.log(`錯誤堆疊: ${error.stack}`);
  }
  
  Logger.log('========================================');
}

/**
 * 檢查 Log 工作表
 */
function checkLogSheet() {
  checkSheetData('Log');
}

/**
 * 列出所有工作表及其資料量
 */
function listAllSheets() {
  Logger.log('========== 所有工作表 ==========');
  
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    
    Logger.log(`試算表名稱: ${ss.getName()}`);
    Logger.log(`試算表網址: ${ss.getUrl()}`);
    Logger.log(`工作表總數: ${sheets.length}\n`);
    
    sheets.forEach((sheet, idx) => {
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      Logger.log(`${idx + 1}. ${sheet.getName()}`);
      Logger.log(`   列數: ${lastRow}, 欄數: ${lastCol}`);
    });
    
  } catch (error) {
    Logger.log(`✗ 列出工作表失敗: ${error.message}`);
  }
  
  Logger.log('================================');
}

/**
 * 檢查 Excel 轉換後的原始資料
 * 用於診斷為何 buildShiftMap 找不到班別代碼
 */
function debugExcelRawData() {
  Logger.log('========== 檢查 Excel 轉換後的原始資料 ==========');
  
  try {
    const ss = getSpreadsheet();
    const sheets = ss.getSheets();
    
    // 找最新上傳的工作表（通常是暫存的轉換結果）
    Logger.log(`\n共有 ${sheets.length} 個工作表:`);
    sheets.forEach((sheet, idx) => {
      Logger.log(`${idx + 1}. ${sheet.getName()} (${sheet.getLastRow()} 列)`);
    });
    
    // 檢查最後一次上傳的暫存檔案（在 Drive 中）
    Logger.log('\n========== 檢查 Drive 中的暫存檔案 ==========');
    const files = DriveApp.getFiles();
    let tempFiles = [];
    while (files.hasNext() && tempFiles.length < 10) {
      const file = files.next();
      if (file.getName().includes('2026泉威國安班表')) {
        tempFiles.push({
          name: file.getName(),
          id: file.getId(),
          mimeType: file.getMimeType(),
          createdDate: file.getDateCreated()
        });
      }
    }
    
    Logger.log(`找到 ${tempFiles.length} 個相關檔案:`);
    tempFiles.forEach((f, idx) => {
      Logger.log(`${idx + 1}. ${f.name} (${f.mimeType})`);
      Logger.log(`   ID: ${f.id}`);
      Logger.log(`   建立時間: ${f.createdDate}`);
      
      // 如果是 Google Sheets，列出工作表
      if (f.mimeType === 'application/vnd.google-apps.spreadsheet') {
        try {
          const tempSs = SpreadsheetApp.openById(f.id);
          const tempSheets = tempSs.getSheets();
          Logger.log(`   工作表 (${tempSheets.length} 個):`);
          tempSheets.forEach((sheet, sidx) => {
            Logger.log(`     ${sidx + 1}. ${sheet.getName()} (${sheet.getLastRow()} 列)`);
          });
        } catch (e) {
          Logger.log(`   ✗ 無法開啟: ${e.message}`);
        }
      }
    });
    
    // 選擇要檢查的工作表（可改為 '11501' 或其他）
    const targetSheetName = '11501';
    const sheet = ss.getSheetByName(targetSheetName);
    
    if (!sheet) {
      Logger.log(`\n✗ 目標試算表中找不到工作表: ${targetSheetName}`);
      Logger.log(`建議: 檢查上方列出的暫存檔案，看 11501 工作表是否在那裡`);
      return;
    }
    
    Logger.log(`\n檢查工作表: ${targetSheetName}`);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log(`總列數: ${lastRow}, 總欄數: ${lastCol}`);
    
    // 檢查前 30 列的第一欄（班別代碼定義通常在這裡）
    Logger.log('\n前 30 列的第一欄內容（尋找 * 開頭的代碼定義）:');
    for (let i = 1; i <= Math.min(30, lastRow); i++) {
      const cellValue = sheet.getRange(i, 1).getValue();
      const cellText = cellValue ? cellValue.toString() : '';
      
      if (cellText.includes('*') || cellText.includes('A') || cellText.includes('B') || cellText.includes('C') || cellText.includes('O')) {
        Logger.log(`第 ${i} 列 A 欄: "${cellText}" ${cellText.indexOf('*') === 0 ? '← ★ 找到 *' : ''}`);
      } else if (i <= 10) {
        Logger.log(`第 ${i} 列 A 欄: "${cellText}"`);
      }
    }
    
    // 檢查 20-30 列範圍（根據本地測試，代碼定義在第 21-26 列）
    Logger.log('\n重點檢查第 20-30 列:');
    for (let i = 20; i <= Math.min(30, lastRow); i++) {
      const cellValue = sheet.getRange(i, 1).getValue();
      const cellText = cellValue ? cellValue.toString() : '';
      Logger.log(`第 ${i} 列 A 欄: "${cellText}" (類型: ${typeof cellValue})`);
    }
    
  } catch (error) {
    Logger.log(`✗ 檢查失敗: ${error.message}`);
    Logger.log(`錯誤堆疊: ${error.stack}`);
  }
  
  Logger.log('================================================');
}
