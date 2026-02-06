/**
 * 薪資計算系統 - Excel 解析器
 * 功能：解析 Excel 檔案並提取指定工作表的資料
 */

/**
 * 解析 Excel 檔案
 * @param {string} base64Data - Base64 編碼的 Excel 檔案內容
 * @param {string} fileName - 檔案名稱
 * @param {string} targetSheetName - 目標工作表名稱（例如：'11501'）
 * @return {Object} 解析結果 {success, data, error}
 */
function parseExcel(base64Data, fileName, targetSheetName) {
  try {
    logToSheet(`開始解析 Excel: ${fileName}, 目標工作表: ${targetSheetName}`, 'INFO');
    
    // 將 Base64 轉為 Blob
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileName
    );
    
    // 將 Blob 暫存到 Drive（GAS 需要透過 Drive 來讀取 Excel）
    const tempFile = DriveApp.createFile(blob);
    const fileId = tempFile.getId();
    
    try {
      // 使用 Drive API 或 Sheets API 轉換 Excel
      // 注意：GAS 無法直接讀取 Excel 的特定 sheet，需要先轉換為 Google Sheets
      const convertedSheet = convertExcelToSheets(fileId);
      
      // 讀取指定工作表的資料
      const data = readSheetData(convertedSheet, targetSheetName);
      
      // 刪除暫存檔案
      DriveApp.getFileById(fileId).setTrashed(true);
      DriveApp.getFileById(convertedSheet.getId()).setTrashed(true);
      
      logToSheet(`Excel 解析成功，共 ${data.length} 列資料`, 'INFO');
      
      return {
        success: true,
        data: data,
        sheetName: targetSheetName,
        rowCount: data.length
      };
      
    } catch (error) {
      // 清理暫存檔案
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (e) {
        Logger.log('清理暫存檔案失敗: ' + e.message);
      }
      throw error;
    }
    
  } catch (error) {
    const errorMsg = `Excel 解析失敗: ${error.message}`;
    logToSheet(errorMsg, 'ERROR');
    return {
      success: false,
      error: errorMsg
    };
  }
}

/**
 * 將 Excel 檔案轉換為 Google Sheets
 * @param {string} fileId - Drive 檔案 ID
 * @return {Spreadsheet} Google Sheets 物件
 */
function convertExcelToSheets(fileId) {
  const file = DriveApp.getFileById(fileId);
  
  // 使用 Drive API 轉換
  const resource = {
    title: file.getName(),
    mimeType: MimeType.GOOGLE_SHEETS
  };
  
  // 注意：這裡使用 Advanced Drive Service
  // 需要在 appsscript.json 中啟用 Drive API
  try {
    const options = {
      supportsAllDrives: true
    };
    
    // 使用 Drive.Files.copy 轉換
    const convertedFile = Drive.Files.copy(resource, fileId, options);
    return SpreadsheetApp.openById(convertedFile.id);
    
  } catch (error) {
    // 如果 Drive API 未啟用，使用替代方案
    Logger.log('Drive API 未啟用，使用替代方案');
    return convertExcelAlternative(fileId);
  }
}

/**
 * 替代方案：手動讀取 Excel（較慢但不需要 Drive API）
 */
function convertExcelAlternative(fileId) {
  // 這是一個簡化版本，實際使用時可能需要額外的函式庫
  // 或者請使用者直接上傳 Google Sheets 格式
  throw new Error('需要啟用 Drive API 才能解析 Excel 檔案');
}

/**
 * 讀取 Google Sheets 中指定工作表的資料
 * @param {Spreadsheet} spreadsheet - Google Sheets 物件
 * @param {string} sheetName - 工作表名稱
 * @return {Array} 二維陣列資料
 */
function readSheetData(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  
  if (!sheet) {
    const availableSheets = spreadsheet.getSheets().map(s => s.getName()).join(', ');
    throw new Error(`找不到工作表 "${sheetName}"。可用的工作表: ${availableSheets}`);
  }
  
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow === 0 || lastCol === 0) {
    logToSheet(`工作表 "${sheetName}" 是空的`, 'WARNING');
    return [];
  }
  
  // 讀取所有資料
  const data = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  
  Logger.log(`讀取工作表 "${sheetName}": ${lastRow} 列 x ${lastCol} 欄`);
  
  return data;
}

/**
 * 列出 Excel 中所有工作表名稱
 * @param {string} base64Data - Base64 編碼的 Excel 檔案內容
 * @param {string} fileName - 檔案名稱
 * @return {Array} 工作表名稱陣列
 */
function listExcelSheets(base64Data, fileName) {
  try {
    const blob = Utilities.newBlob(
      Utilities.base64Decode(base64Data),
      'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      fileName
    );
    
    const tempFile = DriveApp.createFile(blob);
    const fileId = tempFile.getId();
    
    try {
      const convertedSheet = convertExcelToSheets(fileId);
      const sheetNames = convertedSheet.getSheets().map(sheet => sheet.getName());
      
      // 清理暫存檔案
      DriveApp.getFileById(fileId).setTrashed(true);
      DriveApp.getFileById(convertedSheet.getId()).setTrashed(true);
      
      return sheetNames;
      
    } catch (error) {
      try {
        DriveApp.getFileById(fileId).setTrashed(true);
      } catch (e) {
        Logger.log('清理暫存檔案失敗: ' + e.message);
      }
      throw error;
    }
    
  } catch (error) {
    logToSheet(`列出工作表失敗: ${error.message}`, 'ERROR');
    return [];
  }
}

/**
 * 驗證 Excel 資料格式
 * @param {Array} data - 二維陣列資料
 * @return {Object} 驗證結果 {valid, errors, warnings}
 */
function validateExcelData(data) {
  const errors = [];
  const warnings = [];
  
  // 檢查是否有資料
  if (!data || data.length === 0) {
    errors.push('檔案中沒有資料');
    return { valid: false, errors, warnings };
  }
  
  // 檢查是否有標題列
  if (data.length < 2) {
    warnings.push('只有一列資料，可能缺少標題列或資料列');
  }
  
  // 檢查每列的欄位數量是否一致
  const firstRowLength = data[0].length;
  for (let i = 1; i < data.length; i++) {
    if (data[i].length !== firstRowLength) {
      warnings.push(`第 ${i + 1} 列的欄位數量 (${data[i].length}) 與第 1 列 (${firstRowLength}) 不一致`);
    }
  }
  
  // 檢查是否有完全空白的列
  for (let i = 0; i < data.length; i++) {
    const isEmpty = data[i].every(cell => cell === '' || cell === null || cell === undefined);
    if (isEmpty) {
      warnings.push(`第 ${i + 1} 列是空白列`);
    }
  }
  
  return {
    valid: errors.length === 0,
    errors,
    warnings,
    rowCount: data.length,
    columnCount: firstRowLength
  };
}
