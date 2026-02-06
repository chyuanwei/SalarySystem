/**
 * 薪資計算系統 - Excel 解析器（泉威國安班表專用）
 * 功能：解析泉威國安班表格式（代碼式班別、橫向日期）
 */

/**
 * 解析 Excel 檔案
 * @param {string} base64Data - Base64 編碼的 Excel 檔案內容
 * @param {string} fileName - 檔案名稱
 * @param {string} targetSheetName - 目標工作表名稱（例如：'11501'）
 * @return {Object} 解析結果 {success, data, shiftMap, error}
 */
function parseExcel(base64Data, fileName, targetSheetName) {
  try {
    logDebug(`開始解析 Excel: ${fileName}`, {
      fileName: fileName,
      targetSheetName: targetSheetName
    });
    
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
      const convertedSheet = convertExcelToSheets(fileId);
      
      // 讀取指定工作表的資料
      const data = readSheetData(convertedSheet, targetSheetName);
      
      // 刪除暫存檔案
      DriveApp.getFileById(fileId).setTrashed(true);
      DriveApp.getFileById(convertedSheet.getId()).setTrashed(true);
      
      // 使用泉威國安班表解析器
      const parsedResult = parseQuanWeiSchedule(data, targetSheetName);
      
      logInfo(`Excel 解析成功`, {
        fileName: fileName,
        sheetName: targetSheetName,
        totalRecords: parsedResult.records.length,
        totalEmployees: parsedResult.totalEmployees,
        shiftCodes: Object.keys(parsedResult.shiftMap)
      });
      
      return {
        success: true,
        data: parsedResult.records,
        shiftMap: parsedResult.shiftMap,
        config: parsedResult.config,
        totalEmployees: parsedResult.totalEmployees,
        rowCount: parsedResult.records.length
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
    logError(errorMsg, {
      fileName: fileName,
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
  
  // 記錄原始欄位結構（用於分析和映射）
  if (data.length > 0) {
    const headers = data[0];
    logOperation(`分析 Excel 欄位結構`, {
      sheetName: sheetName,
      columnCount: headers.length,
      headers: headers,
      sampleData: data.length > 1 ? data.slice(1, Math.min(4, data.length)) : []
    });
  }
  
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
    logError(`列出工作表失敗: ${error.message}`, {
      fileName: fileName,
      error: error.toString()
    });
    return [];
  }
}

/**
 * 泉威國安班表解析器 - 主函數
 * 功能：解析代碼式班表（橫向日期、代碼定義）
 * @param {Array} data - 二維陣列資料
 * @param {string} sheetName - 工作表名稱
 * @return {Object} 解析結果 {records, shiftMap, config, totalEmployees}
 */
function parseQuanWeiSchedule(data, sheetName) {
  try {
    logOperation(`開始解析泉威國安班表: ${sheetName}`);
    
    // 1. 動態建立代碼字典
    const shiftMap = buildShiftMap(data);
    logDebug('找到的班別代碼', { shiftMap: shiftMap });
    
    // 2. 定位資料區塊
    const config = locateDataStructure(data);
    if (!config) {
      throw new Error('找不到日期列（1, 2, 3...），請確認工作表格式是否正確');
    }
    logDebug('資料結構配置', { config: config });
    
    // 3. 取得年月份（從 A2 取得，例如 2026/01）
    let yearMonth = data[1] && data[1][0] ? data[1][0].toString().trim() : '';
    if (!yearMonth.includes('/')) {
      yearMonth = '2026/01'; // 預設值
      logWarning('無法取得年月份，使用預設值', { yearMonth: yearMonth });
    }
    logDebug('年月份', { yearMonth: yearMonth });
    
    const finalResults = [];
    let processedEmployees = 0;
    
    // 4. 主解析迴圈
    logDebug(`開始從第 ${config.startRow + 1} 列解析員工排班`);
    
    for (let i = config.startRow; i < data.length; i++) {
      const empName = data[i] && data[i][0] ? data[i][0].toString().trim() : '';
      
      // 終止條件：遇到空列或統計標籤
      if (!empName || ['上班人數', '合計', '備註', 'P.T', '閉店評論'].some(k => empName.includes(k))) {
        if (i > config.startRow) break;
        continue;
      }
      
      processedEmployees++;
      let shiftCount = 0;
      
      // 遍歷所有日期欄位
      for (let j = config.dateCol; j < config.maxCol && j < (data[i] ? data[i].length : 0); j++) {
        const dayNum = data[config.dateRow] ? data[config.dateRow][j] : '';
        const cellValue = data[i][j] ? data[i][j].toString().trim() : '';
        
        // 跳過無效資料
        if (!dayNum || isNaN(parseInt(dayNum)) || !cellValue || cellValue.toLowerCase() === 'nan') {
          continue;
        }
        
        let shiftInfo = { start: '', end: '', hours: '' };
        
        // A. 先查代碼字典
        if (shiftMap[cellValue]) {
          shiftInfo = shiftMap[cellValue];
        }
        // B. 若字典查不到，檢查是否為手寫時間（如 18:00-20:30）
        else if (cellValue.includes('-')) {
          shiftInfo = parseTimeRange(cellValue);
        }
        // C. 其他情況跳過
        else {
          continue;
        }
        
        if (shiftInfo.start) {
          finalResults.push([
            empName,
            formatScheduleDate(yearMonth, dayNum),
            shiftInfo.start,
            shiftInfo.end,
            shiftInfo.hours
          ]);
          shiftCount++;
        }
      }
      
      logDebug(`處理員工: ${empName}`, { shiftCount: shiftCount });
    }
    
    logOperation(`泉威國安班表解析完成`, {
      sheetName: sheetName,
      totalRecords: finalResults.length,
      totalEmployees: processedEmployees,
      shiftCodes: Object.keys(shiftMap).join(', ')
    });
    
    return {
      records: finalResults,
      shiftMap: shiftMap,
      config: config,
      totalEmployees: processedEmployees
    };
    
  } catch (error) {
    logError('泉威國安班表解析失敗', {
      sheetName: sheetName,
      error: error.toString()
    });
    throw error;
  }
}

/**
 * 定位資料結構：自動搜尋日期 1 所在的座標
 * @param {Array} data - 二維陣列資料
 * @return {Object} 資料結構配置 {dateRow, dateCol, startRow, maxCol}
 */
function locateDataStructure(data) {
  for (let r = 0; r < 10 && r < data.length; r++) {
    for (let c = 1; c < 15 && c < (data[r] ? data[r].length : 0); c++) {
      if (data[r][c] == 1) {
        return {
          dateRow: r,
          dateCol: c,
          startRow: r + 2, // 員工姓名通常在日期列下方 2 列
          maxCol: data[r].length
        };
      }
    }
  }
  return null;
}

/**
 * 動態掃描代碼定義區（* A 10:00-17:00）
 * @param {Array} data - 二維陣列資料
 * @return {Object} 代碼字典 {A: {start, end, hours}, ...}
 */
function buildShiftMap(data) {
  const map = {};
  data.forEach((row, rowIdx) => {
    // 檢查第一個儲存格是否以 * 開頭
    const firstCell = row[0] ? row[0].toString().trim() : '';
    if (firstCell.indexOf('*') === 0) {
      // 從第一個儲存格提取代碼和時間範圍
      // 格式：* A 10:00-17:00 或 * A1 10:00-15:00
      const parts = firstCell.split(/\s+/);
      if (parts.length >= 3) {
        const code = parts[1]; // A, A1, B, etc.
        const range = parts[2]; // 10:00-17:00
        if (code && range && range.indexOf('-') > 0) {
          map[code] = parseTimeRange(range);
          logDebug(`代碼定義: ${code} → ${range}`, { row: rowIdx + 1 });
        }
      }
    }
  });
  return map;
}

/**
 * 解析時間範圍並計算時數
 * @param {string} str - 時間範圍字串（例如：10:00-17:00）
 * @return {Object} {start, end, hours}
 */
function parseTimeRange(str) {
  try {
    const parts = str.split('-');
    const s = formatTimeString(parts[0]);
    const e = formatTimeString(parts[1]);
    const sParts = s.split(':');
    const eParts = e.split(':');
    const startDate = new Date(0, 0, 0, parseInt(sParts[0]), parseInt(sParts[1]));
    const endDate = new Date(0, 0, 0, parseInt(eParts[0]), parseInt(eParts[1]));
    let diff = (endDate - startDate) / 1000 / 60 / 60;
    if (diff < 0) diff += 24; // 跨日處理
    return { start: s, end: e, hours: diff.toFixed(1) };
  } catch (err) {
    return { start: '', end: '', hours: '' };
  }
}

/**
 * 格式化時間字串為 HH:MM
 * @param {string} t - 時間字串
 * @return {string} 格式化後的時間
 */
function formatTimeString(t) {
  t = t.toString().replace(':', '').trim();
  return (t.length === 4) ? t.substring(0, 2) + ':' + t.substring(2, 4) : t;
}

/**
 * 格式化日期為 YYYY/MM/DD
 * @param {string} ym - 年月（例如：2026/01）
 * @param {number} d - 日期數字
 * @return {string} 格式化後的日期
 */
function formatScheduleDate(ym, d) {
  const dStr = (d < 10) ? '0' + d : d.toString();
  return ym + '/' + dStr;
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
  
  // 檢查最小列數
  if (data.length < 5) {
    warnings.push('資料列數過少，可能不是正確的班表格式');
  }
  
  // 檢查是否有完全空白的列
  let emptyRowCount = 0;
  for (let i = 0; i < data.length; i++) {
    const isEmpty = data[i].every(cell => cell === '' || cell === null || cell === undefined);
    if (isEmpty) emptyRowCount++;
  }
  if (emptyRowCount > data.length / 2) {
    warnings.push(`發現過多空白列 (${emptyRowCount} 列)`);
  }
  
  return {
    valid: errors.length === 0,
    errors,
    warnings,
    rowCount: data.length,
    columnCount: data[0] ? data[0].length : 0
  };
}
