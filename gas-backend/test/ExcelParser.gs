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
      // 注意：GAS 無法直接讀取 Excel 的特定 sheet，需要先轉換為 Google Sheets
      const convertedSheet = convertExcelToSheets(fileId);
      
      // 讀取指定工作表的資料
      const data = readSheetData(convertedSheet, targetSheetName);
      
      // 刪除暫存檔案
      DriveApp.getFileById(fileId).setTrashed(true);
      DriveApp.getFileById(convertedSheet.getId()).setTrashed(true);
      
      logInfo(`Excel 解析成功`, {
        fileName: fileName,
        sheetName: targetSheetName,
        rowCount: data.length
      });
      
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

/**
 * 將上傳的 Excel 資料轉換為目標格式
 * 目標格式：員工姓名 | 排班日期 | 上班時間 | 下班時間 | 工作時數 | 備註
 * @param {Array} sourceData - 原始 Excel 資料（二維陣列）
 * @return {Array} 轉換後的資料
 */
function transformToTargetFormat(sourceData) {
  try {
    if (!sourceData || sourceData.length === 0) {
      throw new Error('來源資料為空');
    }
    
    // 第一列是標題列
    const headers = sourceData[0];
    logDebug('原始欄位標題', { headers: headers });
    
    // 自動偵測欄位映射
    const mapping = detectColumnMapping(headers);
    logOperation('欄位映射結果', { mapping: mapping });
    
    // 建立目標資料（包含標題列）
    const targetData = [];
    targetData.push(['員工姓名', '排班日期', '上班時間', '下班時間', '工作時數', '備註']);
    
    // 轉換每一列資料（跳過標題列）
    for (let i = 1; i < sourceData.length; i++) {
      const sourceRow = sourceData[i];
      
      // 跳過空白列
      const isEmpty = sourceRow.every(cell => cell === '' || cell === null || cell === undefined);
      if (isEmpty) continue;
      
      // 提取欄位值
      const employeeName = getColumnValue(sourceRow, mapping.employeeName);
      const scheduleDate = getColumnValue(sourceRow, mapping.scheduleDate);
      const startTime = getColumnValue(sourceRow, mapping.startTime);
      const endTime = getColumnValue(sourceRow, mapping.endTime);
      const note = getColumnValue(sourceRow, mapping.note);
      
      // 計算工作時數
      const workHours = calculateWorkHours(startTime, endTime);
      
      // 格式化日期（確保為 YYYY-MM-DD 格式）
      const formattedDate = formatDate(scheduleDate);
      
      // 建立目標列
      const targetRow = [
        employeeName,
        formattedDate,
        formatTime(startTime),
        formatTime(endTime),
        workHours,
        note || ''
      ];
      
      targetData.push(targetRow);
    }
    
    logInfo('資料轉換完成', {
      sourceRows: sourceData.length - 1,
      targetRows: targetData.length - 1
    });
    
    return targetData;
    
  } catch (error) {
    logError('資料轉換失敗', { error: error.toString() });
    throw error;
  }
}

/**
 * 自動偵測欄位映射
 * @param {Array} headers - 標題列
 * @return {Object} 欄位索引映射
 */
function detectColumnMapping(headers) {
  const mapping = {
    employeeName: -1,
    scheduleDate: -1,
    startTime: -1,
    endTime: -1,
    note: -1
  };
  
  // 定義可能的欄位名稱（支援多種變體）
  const patterns = {
    employeeName: ['員工姓名', '姓名', '員工', '人員', 'name', 'employee'],
    scheduleDate: ['排班日期', '日期', '工作日期', 'date', 'schedule'],
    startTime: ['上班時間', '開始時間', '起始時間', 'start', 'begin'],
    endTime: ['下班時間', '結束時間', '完成時間', 'end', 'finish'],
    note: ['備註', '說明', '註記', 'note', 'remark', 'comment']
  };
  
  // 遍歷標題列，尋找匹配的欄位
  for (let i = 0; i < headers.length; i++) {
    const header = String(headers[i]).trim().toLowerCase();
    
    for (const [field, keywords] of Object.entries(patterns)) {
      for (const keyword of keywords) {
        if (header.includes(keyword.toLowerCase())) {
          mapping[field] = i;
          break;
        }
      }
    }
  }
  
  // 檢查必要欄位是否都找到
  const missingFields = [];
  if (mapping.employeeName === -1) missingFields.push('員工姓名');
  if (mapping.scheduleDate === -1) missingFields.push('排班日期');
  if (mapping.startTime === -1) missingFields.push('上班時間');
  if (mapping.endTime === -1) missingFields.push('下班時間');
  
  if (missingFields.length > 0) {
    logWarning('部分必要欄位無法自動偵測', {
      missingFields: missingFields,
      headers: headers
    });
  }
  
  return mapping;
}

/**
 * 取得欄位值
 * @param {Array} row - 資料列
 * @param {number} columnIndex - 欄位索引
 * @return {*} 欄位值
 */
function getColumnValue(row, columnIndex) {
  if (columnIndex === -1 || columnIndex >= row.length) {
    return '';
  }
  return row[columnIndex];
}

/**
 * 計算工作時數
 * @param {*} startTime - 上班時間
 * @param {*} endTime - 下班時間
 * @return {number} 工作時數
 */
function calculateWorkHours(startTime, endTime) {
  try {
    // 處理各種時間格式
    const start = parseTime(startTime);
    const end = parseTime(endTime);
    
    if (!start || !end) {
      return 0;
    }
    
    // 計算時數差（單位：小時）
    let hours = (end - start) / (1000 * 60 * 60);
    
    // 如果結束時間小於開始時間，表示跨日
    if (hours < 0) {
      hours += 24;
    }
    
    // 四捨五入到小數點後 2 位
    return Math.round(hours * 100) / 100;
    
  } catch (error) {
    logWarning('計算工作時數失敗', {
      startTime: startTime,
      endTime: endTime,
      error: error.toString()
    });
    return 0;
  }
}

/**
 * 解析時間字串或日期物件
 * @param {*} timeValue - 時間值
 * @return {Date} 日期物件
 */
function parseTime(timeValue) {
  if (!timeValue) return null;
  
  // 如果已經是 Date 物件
  if (timeValue instanceof Date) {
    return timeValue;
  }
  
  // 如果是字串，嘗試解析
  const timeStr = String(timeValue).trim();
  
  // 格式：HH:MM 或 HH:MM:SS
  const timePattern = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/;
  const match = timeStr.match(timePattern);
  
  if (match) {
    const hours = parseInt(match[1]);
    const minutes = parseInt(match[2]);
    const seconds = match[3] ? parseInt(match[3]) : 0;
    
    const date = new Date();
    date.setHours(hours, minutes, seconds, 0);
    return date;
  }
  
  // 嘗試直接解析
  const parsed = new Date(timeValue);
  if (!isNaN(parsed.getTime())) {
    return parsed;
  }
  
  return null;
}

/**
 * 格式化日期為 YYYY-MM-DD
 * @param {*} dateValue - 日期值
 * @return {string} 格式化後的日期字串
 */
function formatDate(dateValue) {
  if (!dateValue) return '';
  
  try {
    let date;
    
    // 如果已經是 Date 物件
    if (dateValue instanceof Date) {
      date = dateValue;
    } else {
      // 嘗試解析字串
      date = new Date(dateValue);
    }
    
    if (isNaN(date.getTime())) {
      return String(dateValue);
    }
    
    const year = date.getFullYear();
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const day = String(date.getDate()).padStart(2, '0');
    
    return `${year}-${month}-${day}`;
    
  } catch (error) {
    logWarning('日期格式化失敗', { dateValue: dateValue, error: error.toString() });
    return String(dateValue);
  }
}

/**
 * 格式化時間為 HH:MM
 * @param {*} timeValue - 時間值
 * @return {string} 格式化後的時間字串
 */
function formatTime(timeValue) {
  if (!timeValue) return '';
  
  try {
    const date = parseTime(timeValue);
    if (!date) return String(timeValue);
    
    const hours = String(date.getHours()).padStart(2, '0');
    const minutes = String(date.getMinutes()).padStart(2, '0');
    
    return `${hours}:${minutes}`;
    
  } catch (error) {
    return String(timeValue);
  }
}
