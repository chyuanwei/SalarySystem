/**
 * 薪資計算系統 - 主程式
 * 測試環境
 */

// 部署標記：每次 clasp push 後可更新此版號，Log 工作表會寫入此值，用以確認程式是否成功部署
var DEPLOY_MARKER_VERSION = 'v0.6.51';

/**
 * 處理 GET 請求
 */
function doGet(e) {
  const action = e.parameter.action;
  
  try {
    switch (action) {
      case 'test':
        return createJsonResponse({ success: true, message: '測試連線成功', environment: 'test' });
      
      case 'status':
        const jobId = e.parameter.jobId;
        return getJobStatus(jobId);
      
      case 'getBranches': {
        const result = readBranchNames();
        return createJsonResponse(result);
      }
      
      case 'getPersonnelByBranch': {
        const branchName = (e.parameter.branch || '').toString().trim() || null;
        const result = getPersonnelByBranch(branchName);
        if (!result.success) {
          return createJsonResponse({ success: false, error: result.error });
        }
        return createJsonResponse({ success: true, names: result.names || [] });
      }

      case 'getPersonnelFromSchedule': {
        const yearMonth = (e.parameter.yearMonth || '').toString().trim() || null;
        const startDate = (e.parameter.startDate || '').toString().trim() || null;
        const endDate = (e.parameter.endDate || '').toString().trim() || null;
        const branchName = (e.parameter.branch || '').toString().trim() || null;
        const result = getPersonnelFromAttendance(yearMonth, startDate, endDate, branchName);
        if (!result.success) {
          return createJsonResponse({ success: false, error: result.error });
        }
        return createJsonResponse({ success: true, names: result.names || [] });
      }
      
      case 'loadSchedule': {
        const config = getConfig();
        const sheetName = config.SHEET_NAMES.SCHEDULE || '班表';
        const yearMonth = (e.parameter.yearMonth || '').toString().trim() || null;
        const dateParam = (e.parameter.date || '').toString().trim() || null;
        const startDate = (e.parameter.startDate || '').toString().trim() || null;
        const endDate = (e.parameter.endDate || '').toString().trim() || null;
        const namesParam = e.parameter.names || '';
        const names = namesParam ? namesParam.split(',').map(function(n) { return n.trim(); }).filter(Boolean) : [];
        const branchName = (e.parameter.branch || '').toString().trim() || null;
        var result;
        if (startDate && startDate.length === 10) {
          result = readScheduleByConditions(yearMonth, startDate, endDate, names, branchName);
          if (!result.success) {
            return createJsonResponse({ success: false, error: result.error });
          }
          var distinctNames = [];
          var seen = {};
          (result.records || []).forEach(function(row) {
            var n = row[0] ? String(row[0]).trim() : '';
            if (n && !seen[n]) { seen[n] = true; distinctNames.push(n); }
          });
          distinctNames.sort();
          return createJsonResponse({
            success: true,
            records: result.records,
            details: { yearMonth: yearMonth || '', date: '', startDate: startDate, endDate: endDate || startDate, names: distinctNames, branch: branchName || '', rowCount: (result.records || []).length }
          });
        }
        result = readScheduleByYearMonth(sheetName, yearMonth, dateParam, names, branchName);
        if (!result.success) {
          return createJsonResponse({ success: false, error: result.error });
        }
        return createJsonResponse({
          success: true,
          records: result.records,
          details: result.details || {}
        });
      }
      
      case 'loadAttendance': {
        const config = getConfig();
        const sheetName = config.SHEET_NAMES.ATTENDANCE || '打卡';
        const yearMonth = (e.parameter.yearMonth || '').toString().trim() || null;
        const dateParam = (e.parameter.date || '').toString().trim() || null;
        const startDate = (e.parameter.startDate || '').toString().trim() || null;
        const endDate = (e.parameter.endDate || '').toString().trim() || null;
        const namesParam = e.parameter.names || '';
        const names = namesParam ? namesParam.split(',').map(function(n) { return n.trim(); }).filter(Boolean) : [];
        const branchName = (e.parameter.branch || '').toString().trim() || null;
        var result;
        if (startDate && startDate.length === 10) {
          result = readAttendanceByConditions(yearMonth, startDate, endDate, names, branchName);
          if (!result.success) {
            return createJsonResponse({ success: false, error: result.error });
          }
          var distinctNames = [];
          var seen = {};
          (result.records || []).forEach(function(row) {
            var n = row[3] ? String(row[3]).trim() : '';
            if (n && !seen[n]) { seen[n] = true; distinctNames.push(n); }
          });
          distinctNames.sort();
          return createJsonResponse({
            success: true,
            records: result.records,
            details: { yearMonth: yearMonth || '', date: '', startDate: startDate, endDate: endDate || startDate, names: distinctNames, branch: branchName || '', rowCount: (result.records || []).length }
          });
        }
        result = readAttendanceByYearMonth(sheetName, yearMonth, dateParam, names, branchName);
        if (!result.success) {
          return createJsonResponse({ success: false, error: result.error });
        }
        return createJsonResponse({
          success: true,
          records: result.records,
          details: result.details || {}
        });
      }
      
      case 'loadCompare': {
        const yearMonth = (e.parameter.yearMonth || '').toString().trim() || null;
        const startDate = (e.parameter.startDate || '').toString().trim() || null;
        const endDate = (e.parameter.endDate || '').toString().trim() || null;
        const namesParam = e.parameter.names || '';
        const names = namesParam ? namesParam.split(',').map(function(n) { return n.trim(); }).filter(Boolean) : [];
        const branchName = (e.parameter.branch || '').toString().trim() || null;
        const result = compareScheduleAttendance(yearMonth, startDate, endDate, names, branchName);
        if (!result.success) {
          return createJsonResponse({ success: false, error: result.error });
        }
        return createJsonResponse({
          success: true,
          items: result.items || []
        });
      }
      
      default:
        return createJsonResponse({ success: false, error: '未知的 action' });
    }
  } catch (error) {
    return createErrorResponse(error);
  }
}

/**
 * 處理 POST 請求
 */
function doPost(e) {
  try {
    if (!e || !e.postData || e.postData.contents === undefined || e.postData.contents === null) {
      return createJsonResponse({ success: false, error: '請求內容為空' });
    }
    var requestData;
    try {
      requestData = JSON.parse(e.postData.contents);
    } catch (parseErr) {
      return createJsonResponse({ success: false, error: '請求格式錯誤（非 JSON）' });
    }
    const action = requestData.action;
    logToSheet('GAS 部署標記: ' + DEPLOY_MARKER_VERSION, 'OPERATION', { action: action });
    logDebug(`收到 POST 請求: ${action}`, {
      action: action,
      timestamp: new Date().toISOString()
    });
    switch (action) {
      case 'upload':
        return handleUpload(requestData);
      
      case 'submitCorrection':
        return handleSubmitCorrection(requestData);
      
      case 'updateScheduleRemark':
        return handleUpdateScheduleRemark(requestData);
      
      case 'updateAttendanceRemark':
        return handleUpdateAttendanceRemark(requestData);
      
      case 'calculate':
        return handleCalculate(requestData);
      
      default:
        logWarning(`收到未知的 action: ${action}`);
        return createJsonResponse({ 
          success: false, 
          error: `未知的 action: ${action}` 
        });
    }
    
  } catch (error) {
    logError(`POST 請求處理失敗: ${error.message}`, {
      error: error.toString(),
      stack: error.stack
    });
    return createErrorResponse(error);
  }
}

/**
 * 處理檔案上傳
 */
function handleUpload(requestData) {
  const startTime = new Date();
  const uploadType = requestData.uploadType || 'schedule';

  try {
    const fileName = requestData.fileName;
    const fileData = requestData.fileData;
    const targetSheetName = requestData.targetSheetName || '11501';
    const targetGoogleSheetTab = requestData.targetGoogleSheetTab || '班表';
    const branchName = requestData.branchName ? String(requestData.branchName).trim() : '';

    // 班表／打卡上傳時分店必選
    if ((uploadType === 'schedule' || uploadType === 'attendance') && !branchName) {
      return createJsonResponse({ success: false, error: '請選擇分店' });
    }

    if (uploadType === 'attendance') {
      return handleAttendanceUpload(requestData, fileName, fileData, branchName, startTime);
    }

    logOperation('開始處理上傳: ' + fileName, {
      fileName: fileName,
      targetSheetName: targetSheetName,
      targetGoogleSheetTab: targetGoogleSheetTab,
      branchName: branchName
    });
    
    // 驗證必要參數
    if (!fileName || !fileData) {
      throw new Error('缺少必要參數: fileName 或 fileData');
    }
    
    logDebug(`檔案大小: ${fileData.length} bytes (Base64)`);
    
    // 解析 Excel 檔案
    const parseResult = parseExcel(fileData, fileName, targetSheetName);
    
    if (!parseResult.success) {
      throw new Error(parseResult.error);
    }
    
    logDebug(`Excel 解析成功，讀取 ${parseResult.rowCount} 列資料`);
    
    // 驗證資料（基本檢查）
    const validation = validateExcelData(parseResult.data);
    
    if (!validation.valid) {
      logError('資料驗證失敗', { errors: validation.errors });
      throw new Error('資料驗證失敗: ' + validation.errors.join(', '));
    }
    
    // 記錄警告
    if (validation.warnings && validation.warnings.length > 0) {
      logWarning(`資料警告`, { warnings: validation.warnings });
    }
    
    // 準備轉換後的資料：每列加 分店、備註、建立時間、修改時間；員工姓名透過人員 sheet 轉成打卡名稱寫入
    logDebug('準備寫入資料');
    var mapping = readPersonnelMapping();
    var scheduleNameToAccount = mapping ? mapping.scheduleNameToAccount : {};
    var accountToAttendanceName = mapping ? mapping.accountToAttendanceName : {};
    var scheduleNowStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    const headerRow = ['員工姓名', '排班日期', '上班時間', '下班時間', '工作時數', '班別', '分店', '備註', '建立時間', '修改時間'];
    const dataWithBranch = parseResult.data.map(function(row) {
      var r = row.slice();
      var scheduleName = r[0] ? String(r[0]).trim() : '';
      var acc = scheduleNameToAccount[scheduleName];
      var attendanceName = acc && accountToAttendanceName[acc] ? accountToAttendanceName[acc] : scheduleName;
      r[0] = attendanceName || scheduleName;
      return r.concat([branchName, '', scheduleNowStr, '']);
    });
    const transformedData = [headerRow].concat(dataWithBranch);
    logDebug(`共 ${parseResult.data.length} 列資料`);
    
    // 以「月份＋分店」覆蓋：先刪除既有該月份該分店資料
    var yearMonth = '';
    var deletedCount = 0;
    if (parseResult.data.length > 0 && parseResult.data[0][1]) {
      yearMonth = extractYearMonth(String(parseResult.data[0][1]));
    }
    if (yearMonth) {
      var removeResult = removeScheduleRowsByYearMonthAndBranch(targetGoogleSheetTab, yearMonth, branchName);
      deletedCount = removeResult.deletedCount || 0;
      if (deletedCount > 0) {
        logInfo('班表覆蓋：已刪除舊資料 ' + deletedCount + ' 筆', { yearMonth: yearMonth, branch: branchName });
      }
    }
    
    // 寫入 Google Sheets（附加模式）
    const writeResult = appendToSheet(transformedData, targetGoogleSheetTab);
    
    if (!writeResult.success) {
      throw new Error(writeResult.error);
    }
    
    // 計算處理時間
    const endTime = new Date();
    const processTime = (endTime - startTime) / 1000; // 秒
    
    // 建立處理記錄
    createProcessRecord({
      fileName: fileName,
      sourceSheet: targetSheetName,
      targetSheet: targetGoogleSheetTab,
      rowCount: writeResult.rowCount,
      status: 'SUCCESS',
      message: '檔案上傳並處理成功'
    });
    
    logOperation(`檔案處理成功: ${fileName}`, {
      fileName: fileName,
      rowCount: writeResult.rowCount,
      totalEmployees: parseResult.totalEmployees,
      columnCount: writeResult.columnCount,
      processTime: `${processTime.toFixed(2)} 秒`,
      shiftCodes: Object.keys(parseResult.shiftMap || {}).join(', '),
      skippedCount: writeResult.skippedCount || 0
    });
    
    // 回傳成功訊息（records 含每筆的 row + isDuplicate，供前端顯示重複記號）
    const records = writeResult.allRecordsWithFlag || (writeResult.appendedRows || []).map(function(row) {
      return { row: row, isDuplicate: false };
    });
    return createJsonResponse({
      success: true,
      message: '檔案已成功上傳並處理',
      columns: ['員工姓名', '排班日期', '上班時間', '下班時間', '工作時數', '班別', '分店', '備註'],
      records: records,
      details: {
        fileName: fileName,
        sourceSheet: targetSheetName,
        targetSheet: targetGoogleSheetTab,
        rowCount: writeResult.rowCount,
        parsedRowCount: parseResult.rowCount,
        totalEmployees: parseResult.totalEmployees,
        columnCount: writeResult.columnCount,
        processTime: processTime.toFixed(2),
        shiftCodes: Object.keys(parseResult.shiftMap || {}),
        warnings: validation.warnings || [],
        skippedCount: writeResult.skippedCount || 0,
        deletedCount: deletedCount
      }
    });
    
  } catch (error) {
    const errorMsg = `檔案上傳處理失敗: ${error.message}`;
    logError(errorMsg, {
      fileName: requestData.fileName || 'unknown',
      error: error.toString(),
      stack: error.stack
    });
    
    // 記錄失敗
    createProcessRecord({
      fileName: requestData.fileName || 'unknown',
      sourceSheet: requestData.targetSheetName || 'unknown',
      targetSheet: requestData.targetGoogleSheetTab || 'unknown',
      rowCount: 0,
      status: 'ERROR',
      message: error.message
    });
    
    return createErrorResponse(error);
  }
}

/**
 * 處理打卡 CSV 上傳
 */
function handleAttendanceUpload(requestData, fileName, fileData, branchName, startTime) {
  try {
    if (!fileName || !fileData) {
      return createJsonResponse({ success: false, error: '缺少必要參數: fileName 或 fileData' });
    }
    logOperation('開始處理打卡上傳: ' + fileName, { fileName: fileName, branchName: branchName });
    
    var parseResult = parseAttendanceCSV(fileData, fileName);
    if (!parseResult.success) {
      return createJsonResponse({ success: false, error: parseResult.error });
    }
    
    var records = parseResult.records || [];
    if (records.length === 0) {
      return createJsonResponse({ success: false, error: 'CSV 無有效資料列' });
    }
    
    var personResult = readPersonnelAccounts();
    if (!personResult.success) {
      return createJsonResponse({ success: false, error: personResult.error || '讀取人員清單失敗' });
    }
    var personAccounts = personResult.accounts || {};
    
    var mappingResult = readBranchLocationMapping();
    if (!mappingResult.success) {
      return createJsonResponse({ success: false, error: mappingResult.error || '讀取分店 mapping 失敗' });
    }
    var locationToName = mappingResult.mapping || {};
    
    var dataRows = [];
    for (var i = 0; i < records.length; i++) {
      var r = records[i];
      var empAccount = r.empAccount || '';
      var location = r.location || '';
      
      if (!personAccounts[empAccount]) {
        return createJsonResponse({ success: false, error: '員工帳號「' + empAccount + '」不存在於人員清單' });
      }
      
      var branchNameFromMapping = locationToName[location];
      if (!branchNameFromMapping) {
        return createJsonResponse({ success: false, error: '打卡地點「' + location + '」無對應分店 mapping' });
      }
      
      if (branchNameFromMapping !== branchName) {
        return createJsonResponse({ success: false, error: '打卡地點與所選分店不符（' + location + ' 對應 ' + branchNameFromMapping + '）' });
      }
      
      dataRows.push([
        branchNameFromMapping,
        r.empNo || '',
        r.empAccount || '',
        r.empName || '',
        r.punchDate || '',
        r.startTime || '',
        r.endTime || '',
        r.workHours || '',
        r.status || '',
        r.remark || '',
        '是',
        '',
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
        ''
      ]);
    }
    
    var config = getConfig();
    var targetSheetName = config.SHEET_NAMES.ATTENDANCE || '打卡';
    var headerRow = ['分店', '員工編號', '員工帳號', '員工姓名', '打卡日期', '上班時間', '下班時間', '工作時數', '狀態', '備註', '是否有效', '校正備註', '建立時間', '校正時間'];
    var dataToWrite = [headerRow].concat(dataRows);
    
    // 以「月份＋分店」覆蓋：先刪除既有該月份該分店資料
    var yearMonth = '';
    var deletedCount = 0;
    if (dataRows.length > 0 && dataRows[0][4]) {
      yearMonth = extractYearMonth(String(dataRows[0][4]));
    }
    if (yearMonth) {
      var removeResult = removeAttendanceRowsByYearMonthAndBranch(targetSheetName, yearMonth, branchName);
      deletedCount = removeResult.deletedCount || 0;
      if (deletedCount > 0) {
        logInfo('打卡覆蓋：已刪除舊資料 ' + deletedCount + ' 筆', { yearMonth: yearMonth, branch: branchName });
      }
    }
    
    var writeResult = appendAttendanceToSheet(dataToWrite, targetSheetName);
    if (!writeResult.success) {
      return createJsonResponse({ success: false, error: writeResult.error || '寫入打卡 sheet 失敗' });
    }
    
    var endTime = new Date();
    var processTime = ((endTime - startTime) / 1000).toFixed(2);
    
    createProcessRecord({
      fileName: fileName,
      sourceSheet: 'CSV',
      targetSheet: targetSheetName,
      rowCount: writeResult.rowCount,
      status: 'SUCCESS',
      message: '打卡 CSV 上傳成功'
    });
    
    logOperation('打卡上傳成功: ' + fileName, {
      fileName: fileName,
      rowCount: writeResult.rowCount,
      skippedCount: writeResult.skippedCount || 0,
      processTime: processTime + ' 秒'
    });
    
    var recordsForFrontend = (writeResult.allRecordsWithFlag || []).map(function(item) {
      return { row: item.row, isDuplicate: item.isDuplicate };
    });
    
    return createJsonResponse({
      success: true,
      message: '打卡檔案已成功上傳並處理',
      columns: headerRow,
      records: recordsForFrontend,
      details: {
        fileName: fileName,
        targetSheet: targetSheetName,
        rowCount: writeResult.rowCount,
        parsedRowCount: records.length,
        columnCount: headerRow.length,
        processTime: processTime,
        skippedCount: writeResult.skippedCount || 0,
        deletedCount: deletedCount
      }
    });
    
  } catch (error) {
    logError('打卡上傳處理失敗: ' + error.message, {
      fileName: fileName || 'unknown',
      error: error.toString(),
      stack: error.stack
    });
    createProcessRecord({
      fileName: fileName || 'unknown',
      sourceSheet: 'CSV',
      targetSheet: '打卡',
      rowCount: 0,
      status: 'ERROR',
      message: error.message
    });
    return createJsonResponse({ success: false, error: '打卡上傳處理失敗: ' + error.message });
  }
}

/**
 * 處理校正送出
 */
function handleSubmitCorrection(requestData) {
  try {
    logToSheet('校正送出-前端傳入內容', 'OPERATION', requestData);
    var branch = (requestData.branch || '').toString().trim();
    var empAccount = (requestData.empAccount || '').toString().trim();
    var name = (requestData.name || '').toString().trim();
    var date = (requestData.date || '').toString().trim();
    if (date && typeof normalizeDateToDash === 'function') date = normalizeDateToDash(date);
    var scheduleStart = (requestData.scheduleStart || '').toString().trim();
    var scheduleEnd = (requestData.scheduleEnd || '').toString().trim();
    var scheduleHours = (requestData.scheduleHours || '').toString().trim();
    var attendanceStart = (requestData.attendanceStart || '').toString().trim();
    var attendanceEnd = (requestData.attendanceEnd || '').toString().trim();
    var attendanceHours = (requestData.attendanceHours || '').toString().trim();
    var attendanceStatus = (requestData.attendanceStatus || '').toString().trim();
    var correctedStart = (requestData.correctedStart || '').toString().trim();
    var correctedEnd = (requestData.correctedEnd || '').toString().trim();
    var remark = (requestData.remark || requestData.correctionRemark || '').toString().trim();
    if (!branch || !date || !correctedStart || !correctedEnd) {
      return createJsonResponse({ success: false, error: '分店、日期、校正上班時間、校正下班時間為必填' });
    }
    var data = {
      branch: branch,
      empAccount: empAccount,
      name: name,
      date: date,
      scheduleStart: scheduleStart,
      scheduleEnd: scheduleEnd,
      scheduleHours: scheduleHours,
      attendanceStart: attendanceStart,
      attendanceEnd: attendanceEnd,
      attendanceHours: attendanceHours,
      attendanceStatus: attendanceStatus,
      correctedStart: correctedStart,
      correctedEnd: correctedEnd,
      remark: remark,
      correctionRemark: remark
    };
    var result = writeCorrectionToAttendance(data);
    if (!result.success) {
      return createJsonResponse({ success: false, error: result.error || '校正寫回打卡失敗' });
    }
    writeCorrection(data);
    return createJsonResponse({ success: true, message: '校正紀錄已送出（已寫回打卡）' });
  } catch (error) {
    logError('校正送出失敗: ' + error.message, { error: error.toString() });
    return createJsonResponse({ success: false, error: '校正送出失敗: ' + error.message });
  }
}

/**
 * 更新班表備註
 */
function handleUpdateScheduleRemark(requestData) {
  try {
    var branch = (requestData.branch || '').toString().trim();
    var name = (requestData.name || '').toString().trim();
    var date = (requestData.date || '').toString().trim();
    var start = (requestData.start || '').toString().trim();
    var end = (requestData.end || '').toString().trim();
    var remark = (requestData.remark || '').toString().trim();
    if (!branch || !name || !date) {
      return createJsonResponse({ success: false, error: '分店、姓名、日期為必填' });
    }
    var result = updateScheduleRemark(branch, name, date, start, end, remark);
    return createJsonResponse(result.success ? { success: true, message: '備註已更新' } : { success: false, error: result.error });
  } catch (err) {
    return createJsonResponse({ success: false, error: '更新班表備註失敗: ' + err.message });
  }
}

/**
 * 更新打卡備註
 */
function handleUpdateAttendanceRemark(requestData) {
  try {
    var branch = (requestData.branch || '').toString().trim();
    var empAccount = (requestData.empAccount || '').toString().trim();
    var date = (requestData.date || '').toString().trim();
    var start = (requestData.start || '').toString().trim();
    var end = (requestData.end || '').toString().trim();
    var remark = (requestData.remark || '').toString().trim();
    if (!branch || !empAccount || !date) {
      return createJsonResponse({ success: false, error: '分店、員工帳號、日期為必填' });
    }
    var result = updateAttendanceRemark(branch, empAccount, date, start, end, remark);
    return createJsonResponse(result.success ? { success: true, message: '備註已更新' } : { success: false, error: result.error });
  } catch (err) {
    return createJsonResponse({ success: false, error: '更新打卡備註失敗: ' + err.message });
  }
}

/**
 * 處理薪資計算（未來實作）
 */
function handleCalculate(requestData) {
  return createJsonResponse({
    success: false,
    error: '薪資計算功能尚未實作'
  });
}

/**
 * 取得任務狀態（未來實作）
 */
function getJobStatus(jobId) {
  return createJsonResponse({
    success: true,
    jobId: jobId,
    status: 'completed',
    message: '狀態查詢功能尚未完整實作'
  });
}

/**
 * 建立 JSON 回應
 */
function createJsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  if (typeof output.setHeader === 'function') {
    output.setHeader('Access-Control-Allow-Origin', '*');
    output.setHeader('Access-Control-Allow-Methods', 'POST, GET, OPTIONS');
    output.setHeader('Access-Control-Allow-Headers', 'Content-Type');
  }
  return output;
}

/**
 * 建立錯誤回應
 */
function createErrorResponse(error) {
  return createJsonResponse({
    success: false,
    error: error.message || String(error),
    timestamp: new Date().toISOString()
  });
}

/**
 * 測試函數 - 可在 GAS 編輯器中直接執行
 */
function testUpload() {
  // 這是一個測試函數，可以用來測試各個功能
  const config = getConfig();
  Logger.log('環境設定:');
  Logger.log(config);
  
  // 測試 Google Sheets 連線
  try {
    const ss = getSpreadsheet();
    Logger.log('Google Sheets 連線成功: ' + ss.getName());
  } catch (error) {
    Logger.log('Google Sheets 連線失敗: ' + error.message);
  }
  
  // 測試寫入記錄
  logToSheet('測試記錄', 'INFO');
  Logger.log('測試完成');
}
