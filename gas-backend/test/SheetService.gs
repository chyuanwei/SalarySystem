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
 * 移除班表 sheet 中符合「年月＋分店」的資料列（覆蓋前刪除舊資料）
 * 班表欄位：員工姓名,排班日期,上班,下班,時數,班別,分店（date=col1, branch=col6）
 * @param {string} sheetName - 工作表名稱（班表）
 * @param {string} yearMonth - 年月 YYYYMM（如 202601）
 * @param {string} branchName - 分店名稱
 * @return {Object} { success, deletedCount }
 */
function removeScheduleRowsByYearMonthAndBranch(sheetName, yearMonth, branchName) {
  try {
    if (!yearMonth || yearMonth.length !== 6 || !branchName) {
      return { success: true, deletedCount: 0 };
    }
    var sheet = getOrCreateSheet(sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, deletedCount: 0 };
    var prefix1 = yearMonth.substring(0, 4) + '/' + yearMonth.substring(4, 6);
    var prefix2 = yearMonth.substring(0, 4) + '-' + yearMonth.substring(4, 6);
    var toDelete = [];
    var data = sheet.getRange(2, 1, lastRow, 7).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var dateVal = row[1];
      var branchVal = (row[6] !== undefined && row[6] !== null) ? String(row[6]).trim() : '';
      if (branchVal !== branchName) continue;
      var dateStr = dateVal ? String(dateVal).trim() : '';
      if (!dateStr) continue;
      var match = dateStr.indexOf(prefix1) === 0 || dateStr.indexOf(prefix2) === 0;
      if (match) toDelete.push(i + 2);
    }
    for (var j = toDelete.length - 1; j >= 0; j--) {
      sheet.deleteRow(toDelete[j]);
    }
    if (toDelete.length > 0) {
      logInfo('班表覆蓋：已刪除舊資料', { sheetName: sheetName, yearMonth: yearMonth, branch: branchName, deletedCount: toDelete.length });
    }
    return { success: true, deletedCount: toDelete.length };
  } catch (error) {
    logError('移除班表舊資料失敗: ' + error.message, { error: error.toString() });
    return { success: false, deletedCount: 0 };
  }
}

/**
 * 移除打卡 sheet 中符合「年月＋分店」的資料列（覆蓋前刪除舊資料）
 * 打卡欄位：分店,員工編號,員工帳號,員工姓名,打卡日期,上班,下班,時數,狀態（branch=col0, date=col4）
 * @param {string} sheetName - 工作表名稱（打卡）
 * @param {string} yearMonth - 年月 YYYYMM（如 202601）
 * @param {string} branchName - 分店名稱
 * @return {Object} { success, deletedCount }
 */
function removeAttendanceRowsByYearMonthAndBranch(sheetName, yearMonth, branchName) {
  try {
    if (!yearMonth || yearMonth.length !== 6 || !branchName) {
      return { success: true, deletedCount: 0 };
    }
    var sheet = getOrCreateSheet(sheetName);
    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return { success: true, deletedCount: 0 };
    var prefix1 = yearMonth.substring(0, 4) + '/' + yearMonth.substring(4, 6);
    var prefix2 = yearMonth.substring(0, 4) + '-' + yearMonth.substring(4, 6);
    var toDelete = [];
    var data = sheet.getRange(2, 1, lastRow, 9).getValues();
    for (var i = 0; i < data.length; i++) {
      var row = data[i];
      var branchVal = (row[0] !== undefined && row[0] !== null) ? String(row[0]).trim() : '';
      if (branchVal !== branchName) continue;
      var dateVal = row[4];
      var dateStr = dateVal ? String(dateVal).trim() : '';
      if (!dateStr) continue;
      var match = dateStr.indexOf(prefix1) === 0 || dateStr.indexOf(prefix2) === 0;
      if (match) toDelete.push(i + 2);
    }
    for (var j = toDelete.length - 1; j >= 0; j--) {
      sheet.deleteRow(toDelete[j]);
    }
    if (toDelete.length > 0) {
      logInfo('打卡覆蓋：已刪除舊資料', { sheetName: sheetName, yearMonth: yearMonth, branch: branchName, deletedCount: toDelete.length });
    }
    return { success: true, deletedCount: toDelete.length };
  } catch (error) {
    logError('移除打卡舊資料失敗: ' + error.message, { error: error.toString() });
    return { success: false, deletedCount: 0 };
  }
}

/**
 * 從日期字串或資料列取得 YYYYMM
 * @param {string} dateStr - YYYY/MM/DD 或 YYYY-MM-DD
 * @return {string} YYYYMM 或 ''
 */
function extractYearMonth(dateStr) {
  if (!dateStr || typeof dateStr !== 'string') return '';
  var s = dateStr.trim().replace(/-/g, '/');
  var parts = s.split('/');
  if (parts.length >= 2) {
    var y = parts[0];
    var m = parts[1];
    if (y && m && y.length === 4 && m.length >= 1) {
      return y + (m.length === 1 ? '0' + m : m).substring(0, 2);
    }
  }
  return '';
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
      const headerRow = data[0] || [];
      const dataRows = data.slice(1); // 去除標題列
      
      // 建立現有資料的去重 key（忽略標題列）
      const existingKeys = new Set();
      if (!isFirstWrite) {
        const lastRow = sheet.getLastRow();
        const lastCol = sheet.getLastColumn();
        if (lastRow >= 2 && lastCol > 0) {
          const colCountForKey = Math.min(7, lastCol);
          const existingData = sheet.getRange(2, 1, lastRow - 1, colCountForKey).getValues();
          existingData.forEach(row => {
            const key = buildDedupKey(row);
            existingKeys.add(key);
          });
        }
      }
      
      // 過濾重複資料（同批內也去重），並建立每筆的 isDuplicate 旗標供前端顯示
      const newRows = [];
      const allRecordsWithFlag = [];
      dataRows.forEach(row => {
        const key = buildDedupKey(row);
        if (!existingKeys.has(key)) {
          existingKeys.add(key);
          newRows.push(row);
          allRecordsWithFlag.push({ row: row, isDuplicate: false });
        } else {
          allRecordsWithFlag.push({ row: row, isDuplicate: true });
        }
      });
      
      const rowsToWrite = isFirstWrite ? [headerRow].concat(newRows) : newRows;
      if (rowsToWrite.length === 0) {
        logWarning('沒有新資料需要追加');
        return {
          success: true,
          message: '沒有新資料需要追加',
          rowCount: 0,
          columnCount: 0,
          skippedCount: dataRows.length,
          appendedRows: [],
          allRecordsWithFlag: allRecordsWithFlag
        };
      }
      
      const startRow = sheet.getLastRow() + 1;
      const rowCount = rowsToWrite.length;
      const colCount = rowsToWrite[0].length;
      
      sheet.getRange(startRow, 1, rowCount, colCount).setValues(rowsToWrite);
      
      // 如果是第一次寫入，格式化標題列
      if (isFirstWrite && data.length > 0) {
        formatHeaderRow(sheet, colCount);
        autoResizeColumns(sheet, colCount);
      }
      
      const appendedCount = newRows.length;
      const skippedCount = dataRows.length - newRows.length;
      const message = `成功追加 ${appendedCount} 列資料到 "${targetSheetName}"`;
      logInfo(message, {
        sheetName: targetSheetName,
        rowCount: appendedCount,
        columnCount: colCount,
        startRow: startRow,
        skippedCount: skippedCount
      });
      
      return {
        success: true,
        message: message,
        rowCount: appendedCount,
        columnCount: colCount,
        sheetName: targetSheetName,
        skippedCount: skippedCount,
        appendedRows: newRows,
        allRecordsWithFlag: allRecordsWithFlag
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
 * 生成去重 Key：員工姓名 + 排班日期 + 上班時間 + 下班時間 + 班別 + 分店
 * @param {Array} row - 單列資料（7 欄時含分店）
 * @return {string} 去重 Key
 */
function buildDedupKey(row) {
  const name = row[0] ? row[0].toString().trim() : '';
  const date = normalizeDateValue(row[1]);
  const start = normalizeTimeValue(row[2]);
  const end = normalizeTimeValue(row[3]);
  const shift = row[5] ? row[5].toString().trim().toUpperCase() : '';
  const branch = (row[6] !== undefined && row[6] !== null) ? String(row[6]).trim() : '';
  return [name, date, start, end, shift, branch].join('|');
}

/**
 * 將日期欄位統一為 YYYY/MM/DD
 * @param {*} value
 * @return {string}
 */
function normalizeDateValue(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
  const text = value.toString().trim();
  if (!text) return '';
  const parsed = new Date(text);
  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy/MM/dd');
  }
  return text;
}

/**
 * 將時間欄位統一為 HH:mm
 * @param {*} value - Date 物件、數字（日的小數，0.41667=10:00）、或字串
 * @return {string}
 */
function normalizeTimeValue(value) {
  if (value === null || value === undefined || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
  }
  if (typeof value === 'number') {
    var fraction = value >= 1 ? value - Math.floor(value) : value;
    if (fraction >= 0 && fraction < 1) {
      var totalMins = Math.round(fraction * 24 * 60);
      var h = Math.floor(totalMins / 60) % 24;
      var m = totalMins % 60;
      var hStr = h < 10 ? '0' + h : '' + h;
      var mStr = m < 10 ? '0' + m : '' + m;
      return hStr + ':' + mStr;
    }
  }
  const text = value.toString().trim();
  if (!text) return '';
  const parsed = new Date(text);
  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'HH:mm');
  }
  return text;
}

/**
 * 依年月/日期、分店與人員篩選班表資料（AND 關係）
 * @param {string} sheetName - 工作表名稱（預設 班表）
 * @param {string} yearMonth - 年月格式 YYYYMM（如 202601），可與 date 二選一
 * @param {string} date - 單日格式 YYYY-MM-DD（如 2026-01-15），可與 yearMonth 二選一
 * @param {Array<string>} names - 員工姓名陣列，可多選（AND 篩選）
 * @param {string|null} branchName - 分店名稱，可選（null 或空表示不限）
 * @return {Object} { success, records, details }
 */
function readScheduleByYearMonth(sheetName, yearMonth, date, names, branchName) {
  try {
    if ((!yearMonth || yearMonth.length !== 6) && (!date || date.length !== 10)) {
      return { success: false, error: '請提供 yearMonth（YYYYMM）或 date（YYYY-MM-DD）' };
    }

    const allData = readFromSheet(sheetName || '班表');
    if (allData.length < 2) {
      return {
        success: true,
        records: [],
        details: { yearMonth: yearMonth || '', date: date || '', names: [], branch: branchName || '', rowCount: 0 }
      };
    }

    const dataRows = allData.slice(1);
    const dateColIndex = 1;
    const nameColIndex = 0;

    var targetPrefix = '';
    var targetExact = '';
    if (date && date.length === 10) {
      targetExact = date.replace(/-/g, '/');
    } else if (yearMonth && yearMonth.length === 6) {
      targetPrefix = yearMonth.substring(0, 4) + '/' + yearMonth.substring(4, 6);
    }

    const nameSet = (names && Array.isArray(names) && names.length > 0)
      ? new Set(names.map(function(n) { return String(n).trim(); }))
      : null;

    const filteredByDate = [];
    dataRows.forEach(function(row) {
      const dateVal = row[dateColIndex];
      if (!dateVal) return;
      var dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      } else {
        dateStr = dateVal.toString().trim();
      }
      var match = false;
      if (targetExact) {
        match = dateStr === targetExact;
      } else if (targetPrefix) {
        match = dateStr.indexOf(targetPrefix) === 0;
      }
      if (!match) return;
      var normalizedRow = row.slice();
      if (dateVal instanceof Date) {
        normalizedRow[dateColIndex] = dateStr;
      }
      normalizedRow[2] = normalizeTimeValue(row[2]);
      normalizedRow[3] = normalizeTimeValue(row[3]);
      filteredByDate.push(normalizedRow);
    });

    var filteredRows = filteredByDate;
    if (branchName && branchName.length > 0) {
      const branchColIndex = 6;
      filteredRows = filteredRows.filter(function(row) {
        const b = row[branchColIndex] ? String(row[branchColIndex]).trim() : '';
        return b === branchName;
      });
    }
    if (nameSet && nameSet.size > 0) {
      filteredRows = filteredRows.filter(function(row) {
        const n = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
        return nameSet.has(n);
      });
    }

    var distinctNames = [];
    const seen = {};
    filteredByDate.forEach(function(row) {
      const n = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
      if (n && !seen[n]) {
        seen[n] = true;
        distinctNames.push(n);
      }
    });
    distinctNames.sort();

    return {
      success: true,
      records: filteredRows,
      details: {
        yearMonth: yearMonth || '',
        date: date || '',
        names: distinctNames,
        branch: branchName || '',
        rowCount: filteredRows.length
      }
    };

  } catch (error) {
    logError('讀取班表失敗: ' + error.message, {
      sheetName: sheetName,
      yearMonth: yearMonth,
      error: error.toString()
    });
    return { success: false, error: error.message };
  }
}

/**
 * 依年月/日期、分店與人員篩選打卡資料（AND 關係）
 * 打卡 sheet 欄位：分店, 員工編號, 員工帳號, 員工姓名, 打卡日期, 上班時間, 下班時間, 工作時數, 狀態
 * @param {string} sheetName - 工作表名稱（預設 打卡）
 * @param {string} yearMonth - 年月格式 YYYYMM
 * @param {string} date - 單日格式 YYYY-MM-DD
 * @param {Array<string>} names - 員工姓名陣列
 * @param {string|null} branchName - 分店名稱
 * @return {Object} { success, records, details }
 */
function readAttendanceByYearMonth(sheetName, yearMonth, date, names, branchName) {
  try {
    if ((!yearMonth || yearMonth.length !== 6) && (!date || date.length !== 10)) {
      return { success: false, error: '請提供 yearMonth（YYYYMM）或 date（YYYY-MM-DD）' };
    }

    var allData = readFromSheet(sheetName || '打卡');
    if (allData.length < 2) {
      return {
        success: true,
        records: [],
        details: { yearMonth: yearMonth || '', date: date || '', names: [], branch: branchName || '', rowCount: 0 }
      };
    }

    var dataRows = allData.slice(1);
    var dateColIndex = 4;
    var nameColIndex = 3;

    var targetPrefix = '';
    var targetExact = '';
    if (date && date.length === 10) {
      targetExact = date.replace(/-/g, '/');
    } else if (yearMonth && yearMonth.length === 6) {
      targetPrefix = yearMonth.substring(0, 4) + '/' + yearMonth.substring(4, 6);
    }

    var nameSet = (names && Array.isArray(names) && names.length > 0)
      ? {}
      : null;
    if (nameSet) {
      names.forEach(function(n) { nameSet[String(n).trim()] = true; });
    }

    var filteredByDate = [];
    dataRows.forEach(function(row) {
      var dateVal = row[dateColIndex];
      if (!dateVal) return;
      var dateStr = '';
      if (dateVal instanceof Date) {
        dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy/MM/dd');
      } else {
        dateStr = dateVal.toString().trim();
      }
      var match = false;
      if (targetExact) {
        match = dateStr === targetExact;
      } else if (targetPrefix) {
        match = dateStr.indexOf(targetPrefix) === 0;
      }
      if (!match) return;
      var normalizedRow = row.slice();
      if (dateVal instanceof Date) {
        normalizedRow[dateColIndex] = dateStr;
      }
      if (row[5]) normalizedRow[5] = normalizeTimeValue(row[5]);
      if (row[6]) normalizedRow[6] = normalizeTimeValue(row[6]);
      filteredByDate.push(normalizedRow);
    });

    var filteredRows = filteredByDate;
    if (branchName && branchName.length > 0) {
      var branchColIndex = 0;
      filteredRows = filteredRows.filter(function(row) {
        var b = row[branchColIndex] ? String(row[branchColIndex]).trim() : '';
        return b === branchName;
      });
    }
    if (nameSet && Object.keys(nameSet).length > 0) {
      filteredRows = filteredRows.filter(function(row) {
        var n = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
        return nameSet[n];
      });
    }

    var distinctNames = [];
    var seen = {};
    filteredByDate.forEach(function(row) {
      var n = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
      if (n && !seen[n]) {
        seen[n] = true;
        distinctNames.push(n);
      }
    });
    distinctNames.sort();

    return {
      success: true,
      records: filteredRows,
      details: {
        yearMonth: yearMonth || '',
        date: date || '',
        names: distinctNames,
        branch: branchName || '',
        rowCount: filteredRows.length
      }
    };

  } catch (error) {
    logError('讀取打卡失敗: ' + error.message, {
      sheetName: sheetName,
      yearMonth: yearMonth,
      error: error.toString()
    });
    return { success: false, error: error.message };
  }
}

/**
 * 讀取人員 sheet 的員工帳號集合（用於打卡驗證）
 * 人員 sheet 欄位：A=員工帳號, B=班表名稱, C=打卡名稱, D=分店
 * @return {Object} { success, accounts: Set<string> } 或 { success: false, error }
 */
function readPersonnelAccounts() {
  try {
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.PERSONNEL || '人員';
    var allData = readFromSheet(sheetName);
    if (!allData || allData.length < 2) {
      return { success: true, accounts: {} };
    }
    var accounts = {};
    var dataRows = allData.slice(1);
    dataRows.forEach(function(row) {
      var acc = row[0] ? String(row[0]).trim() : '';
      if (acc) accounts[acc] = true;
    });
    return { success: true, accounts: accounts };
  } catch (error) {
    logError('讀取人員清單失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message };
  }
}

/**
 * 讀取分店 sheet 的打卡地點 → 名稱 mapping
 * 分店 sheet 欄位：A=代碼, B=名稱, C=啟用狀態, D=排序, E=打卡地點
 * @return {Object} { success, mapping: { 打卡地點: 名稱 } } 或 { success: false, error }
 */
function readBranchLocationMapping() {
  try {
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.BRANCH || '分店';
    var allData = readFromSheet(sheetName);
    var mapping = {};
    if (!allData || allData.length < 2) {
      return { success: true, mapping: mapping };
    }
    var dataRows = allData.slice(1);
    var nameColIndex = 1;
    var locationColIndex = 4;
    dataRows.forEach(function(row) {
      var name = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
      var loc = row[locationColIndex] ? String(row[locationColIndex]).trim() : '';
      if (loc && name) mapping[loc] = name;
    });
    return { success: true, mapping: mapping };
  } catch (error) {
    logError('讀取分店 mapping 失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message };
  }
}

/**
 * 生成打卡去重 Key：員工帳號 + 打卡日期 + 上班時間 + 下班時間 + 分店
 * @param {Array} row - 打卡列 [分店, 員工編號, 員工帳號, 員工姓名, 打卡日期, 上班時間, 下班時間, 工作時數, 狀態]
 * @return {string}
 */
function buildAttendanceDedupKey(row) {
  var branch = (row[0] !== undefined && row[0] !== null) ? String(row[0]).trim() : '';
  var acc = (row[2] !== undefined && row[2] !== null) ? String(row[2]).trim() : '';
  var date = (row[4] !== undefined && row[4] !== null) ? String(row[4]).trim() : '';
  var start = (row[5] !== undefined && row[5] !== null) ? normalizeTimeValue(row[5]) : '';
  var end = (row[6] !== undefined && row[6] !== null) ? normalizeTimeValue(row[6]) : '';
  return [acc || '', date, start, end, branch].join('|');
}

/**
 * 追加打卡資料到「打卡」sheet（含去重）
 * @param {Array} data - 二維陣列，row1 為標題，row2+ 為資料
 * @param {string} targetSheetName - 目標工作表（打卡）
 * @return {Object} { success, rowCount, skippedCount, appendedRows, allRecordsWithFlag }
 */
function appendAttendanceToSheet(data, targetSheetName) {
  try {
    if (!data || data.length < 2) {
      return { success: true, rowCount: 0, skippedCount: 0, appendedRows: [], allRecordsWithFlag: [] };
    }
    var sheet = getOrCreateSheet(targetSheetName);
    var isFirstWrite = (sheet.getLastRow() === 0);
    var headerRow = data[0];
    var dataRows = data.slice(1);
    
    var existingKeys = {};
    if (!isFirstWrite) {
      var lastRow = sheet.getLastRow();
      var lastCol = sheet.getLastColumn();
      if (lastRow >= 2 && lastCol >= 7) {
        var existingData = sheet.getRange(2, 1, lastRow, Math.min(9, lastCol)).getValues();
        existingData.forEach(function(row) {
          var key = buildAttendanceDedupKey(row);
          existingKeys[key] = true;
        });
      }
    }
    
    var newRows = [];
    var allRecordsWithFlag = [];
    dataRows.forEach(function(row) {
      var key = buildAttendanceDedupKey(row);
      if (!existingKeys[key]) {
        existingKeys[key] = true;
        newRows.push(row);
        allRecordsWithFlag.push({ row: row, isDuplicate: false });
      } else {
        allRecordsWithFlag.push({ row: row, isDuplicate: true });
      }
    });
    
    if (newRows.length === 0) {
      logWarning('打卡：無新資料需追加');
      return {
        success: true,
        rowCount: 0,
        skippedCount: dataRows.length,
        appendedRows: [],
        allRecordsWithFlag: allRecordsWithFlag
      };
    }
    
    var rowsToWrite = isFirstWrite ? [headerRow].concat(newRows) : newRows;
    var startRow = sheet.getLastRow() + 1;
    var rowCount = rowsToWrite.length;
    var colCount = rowsToWrite[0].length;
    
    sheet.getRange(startRow, 1, rowCount, colCount).setValues(rowsToWrite);
    
    if (isFirstWrite && data.length > 0) {
      formatHeaderRow(sheet, colCount);
      autoResizeColumns(sheet, colCount);
    }
    
    logInfo('打卡資料追加成功', {
      sheetName: targetSheetName,
      appendedCount: newRows.length,
      skippedCount: dataRows.length - newRows.length
    });
    
    return {
      success: true,
      rowCount: newRows.length,
      skippedCount: dataRows.length - newRows.length,
      appendedRows: newRows,
      allRecordsWithFlag: allRecordsWithFlag
    };
  } catch (error) {
    logError('打卡資料追加失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message };
  }
}

/**
 * 讀取分店清單（從「分店」工作表，欄位：代碼、名稱、啟用狀態、排序）
 * @return {Object} { success, names: [...] } 或 { success: false, error }
 */
function readBranchNames() {
  try {
    const config = getConfig();
    const sheetName = config.SHEET_NAMES.BRANCH || '分店';
    const allData = readFromSheet(sheetName);
    if (!allData || allData.length < 2) {
      return { success: true, names: [] };
    }
    // 結構：A=代碼, B=名稱, C=啟用狀態, D=排序
    const dataRows = allData.slice(1);
    const result = [];
    dataRows.forEach(function(row) {
      const name = row[1] ? String(row[1]).trim() : '';
      const enabled = row[2] ? String(row[2]).trim() : '';
      const sortOrder = row[3] !== undefined && row[3] !== null ? (typeof row[3] === 'number' ? row[3] : parseInt(row[3], 10) || 0) : 0;
      if (name && (enabled === '是' || enabled === 'Y' || enabled === '1' || enabled === '')) {
        result.push({ name: name, sortOrder: isNaN(sortOrder) ? 0 : sortOrder });
      }
    });
    result.sort(function(a, b) { return a.sortOrder - b.sortOrder; });
    const names = result.map(function(r) { return r.name; });
    return { success: true, names: names };
  } catch (error) {
    logError('讀取分店清單失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message };
  }
}

/**
 * 讀取工作表的所有資料（若工作表不存在則回傳空陣列，不建立新工作表）
 * @param {string} sheetName - 工作表名稱
 * @return {Array} 二維陣列資料
 */
function readFromSheet(sheetName) {
  try {
    const sheet = getSheet(sheetName);
    if (!sheet) {
      logDebug(`工作表不存在: ${sheetName}，回傳空陣列`, { sheetName: sheetName });
      return [];
    }
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
