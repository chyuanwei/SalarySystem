/**
 * 薪資計算系統 - Google Sheets 服務
 * 功能：將資料寫入 Google Sheets
 */

/**
 * 將資料寫入指定的 Google Sheets 工作表
 * @param {Array} data - 二維陣列資料
 * @param {string} targetSheetName - 目標工作表名稱
 * @return {Object} 寫入結果 {success, message, rowCount}
 */
function writeToSheet(data, targetSheetName) {
  try {
    logDebug(`開始寫入資料到工作表: ${targetSheetName}`, {
      targetSheetName: targetSheetName,
      rowCount: data ? data.length : 0
    });
    
    if (!data || data.length === 0) {
      throw new Error('沒有資料可寫入');
    }
    
    // 取得或建立目標工作表
    const sheet = getOrCreateSheet(targetSheetName);
    
    // 清空現有資料
    clearSheet(sheet);
    
    // 寫入資料
    const rowCount = data.length;
    const colCount = data[0].length;
    
    sheet.getRange(1, 1, rowCount, colCount).setValues(data);
    
    // 格式化標題列
    formatHeaderRow(sheet, colCount);
    
    // 自動調整欄寬
    autoResizeColumns(sheet, colCount);
    
    // 記錄寫入資訊
    const message = `成功寫入資料到 "${targetSheetName}"`;
    logInfo(message, {
      sheetName: targetSheetName,
      rowCount: rowCount,
      columnCount: colCount
    });
    
    return {
      success: true,
      message: message,
      rowCount: rowCount,
      columnCount: colCount,
      sheetName: targetSheetName
    };
    
  } catch (error) {
    const errorMsg = `寫入 Google Sheets 失敗: ${error.message}`;
    logError(errorMsg, {
      targetSheetName: targetSheetName,
      error: error.toString()
    });
    return {
      success: false,
      error: errorMsg
    };
  }
}

/**
 * 清空工作表的所有資料
 * @param {Sheet} sheet - Google Sheets 工作表物件
 */
function clearSheet(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  
  if (lastRow > 0 && lastCol > 0) {
    sheet.getRange(1, 1, lastRow, lastCol).clear();
    Logger.log(`已清空工作表 "${sheet.getName()}"`);
  }
}

/**
 * 格式化標題列
 * @param {Sheet} sheet - Google Sheets 工作表物件
 * @param {number} colCount - 欄位數量
 */
function formatHeaderRow(sheet, colCount) {
  if (sheet.getLastRow() < 1) return;
  
  const headerRange = sheet.getRange(1, 1, 1, colCount);
  
  // 設定背景色
  headerRange.setBackground('#4a86e8');
  
  // 設定文字顏色
  headerRange.setFontColor('#ffffff');
  
  // 設定粗體
  headerRange.setFontWeight('bold');
  
  // 設定水平對齊
  headerRange.setHorizontalAlignment('center');
  
  // 凍結標題列
  sheet.setFrozenRows(1);
  
  Logger.log('標題列格式化完成');
}

/**
 * 自動調整欄寬
 * @param {Sheet} sheet - Google Sheets 工作表物件
 * @param {number} colCount - 欄位數量
 */
function autoResizeColumns(sheet, colCount) {
  try {
    for (let i = 1; i <= colCount; i++) {
      sheet.autoResizeColumn(i);
    }
    Logger.log('欄寬自動調整完成');
  } catch (error) {
    Logger.log('自動調整欄寬失敗: ' + error.message);
  }
}

/**
 * 追加資料到工作表末尾
 * @param {Array} data - 資料（可以是一維陣列或二維陣列）
 * @param {string} targetSheetName - 目標工作表名稱
 * @return {Object} 追加結果 {success, message, rowCount, columnCount}
 */
function appendToSheet(data, targetSheetName) {
  try {
    logDebug(`開始追加資料到工作表: ${targetSheetName}`, {
      targetSheetName: targetSheetName,
      dataType: Array.isArray(data[0]) ? '二維陣列' : '一維陣列'
    });
    
    if (!data || data.length === 0) {
      throw new Error('沒有資料可追加');
    }
    
    const sheet = getOrCreateSheet(targetSheetName);
    const isFirstWrite = (sheet.getLastRow() === 0);
    
    // 判斷是單列資料還是多列資料
    const is2DArray = Array.isArray(data[0]);
    
    if (is2DArray) {
      // 二維陣列（多列資料）
      const dataToAppend = data.slice(isFirstWrite ? 0 : 1); // 如果不是第一次寫入，跳過標題列
      
      if (dataToAppend.length === 0) {
        logWarning('跳過標題列後沒有資料可追加');
        return {
          success: true,
          message: '沒有新資料需要追加',
          rowCount: 0,
          columnCount: 0
        };
      }
      
      const startRow = sheet.getLastRow() + 1;
      const rowCount = dataToAppend.length;
      const colCount = dataToAppend[0].length;
      
      sheet.getRange(startRow, 1, rowCount, colCount).setValues(dataToAppend);
      
      // 如果是第一次寫入，格式化標題列
      if (isFirstWrite && data.length > 0) {
        formatHeaderRow(sheet, colCount);
        autoResizeColumns(sheet, colCount);
      }
      
      const message = `成功追加 ${rowCount} 列資料到 "${targetSheetName}"`;
      logInfo(message, {
        sheetName: targetSheetName,
        rowCount: rowCount,
        columnCount: colCount,
        startRow: startRow
      });
      
      return {
        success: true,
        message: message,
        rowCount: rowCount,
        columnCount: colCount,
        sheetName: targetSheetName
      };
      
    } else {
      // 一維陣列（單列資料）
      sheet.appendRow(data);
      
      const message = `成功追加 1 列資料到 "${targetSheetName}"`;
      logInfo(message, { sheetName: targetSheetName });
      
      return {
        success: true,
        message: message,
        rowCount: 1,
        columnCount: data.length,
        sheetName: targetSheetName
      };
    }
    
  } catch (error) {
    const errorMsg = `追加資料失敗: ${error.message}`;
    logError(errorMsg, {
      targetSheetName: targetSheetName,
      error: error.toString()
    });
    return {
      success: false,
      error: errorMsg
    };
  }
}

/**
 * 讀取工作表的所有資料
 * @param {string} sheetName - 工作表名稱
 * @return {Array} 二維陣列資料
 */
function readFromSheet(sheetName) {
  try {
    const sheet = getOrCreateSheet(sheetName);
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    
    if (lastRow === 0 || lastCol === 0) {
      return [];
    }
    
    return sheet.getRange(1, 1, lastRow, lastCol).getValues();
    
  } catch (error) {
    logError(`讀取工作表失敗: ${error.message}`, {
      sheetName: sheetName,
      error: error.toString()
    });
    return [];
  }
}

/**
 * 更新特定儲存格的值
 * @param {string} sheetName - 工作表名稱
 * @param {number} row - 列號（從 1 開始）
 * @param {number} col - 欄號（從 1 開始）
 * @param {*} value - 要設定的值
 * @return {Object} 更新結果 {success, message}
 */
function updateCell(sheetName, row, col, value) {
  try {
    const sheet = getOrCreateSheet(sheetName);
    sheet.getRange(row, col).setValue(value);
    
    return {
      success: true,
      message: `成功更新儲存格 ${sheetName}!${columnToLetter(col)}${row}`
    };
    
  } catch (error) {
    return {
      success: false,
      error: `更新儲存格失敗: ${error.message}`
    };
  }
}

/**
 * 將欄號轉換為欄位字母（例如：1 -> A, 27 -> AA）
 * @param {number} column - 欄號（從 1 開始）
 * @return {string} 欄位字母
 */
function columnToLetter(column) {
  let temp;
  let letter = '';
  while (column > 0) {
    temp = (column - 1) % 26;
    letter = String.fromCharCode(temp + 65) + letter;
    column = (column - temp - 1) / 26;
  }
  return letter;
}

/**
 * 建立處理記錄
 * @param {Object} info - 處理資訊
 */
function createProcessRecord(info) {
  try {
    const config = getConfig();
    const sheet = getOrCreateSheet(config.SHEET_NAMES.LOGS);
    
    const record = [
      new Date(),
      info.fileName || '',
      info.sourceSheet || '',
      info.targetSheet || '',
      info.rowCount || 0,
      info.status || 'SUCCESS',
      info.message || ''
    ];
    
    // 如果是第一筆記錄，加入標題
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['時間', '檔案名稱', '來源工作表', '目標工作表', '資料筆數', '狀態', '訊息']);
      formatHeaderRow(sheet, 7);
    }
    
    sheet.appendRow(record);
    
  } catch (error) {
    Logger.log('建立處理記錄失敗: ' + error.message);
  }
}
