/**
 * 薪資計算系統 - 主程式
 * 測試環境
 */

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
      
      case 'loadSchedule': {
        const config = getConfig();
        const sheetName = config.SHEET_NAMES.SCHEDULE || '班表';
        const yearMonth = e.parameter.yearMonth || '';
        const dateParam = e.parameter.date || '';
        const namesParam = e.parameter.names || '';
        const names = namesParam ? namesParam.split(',').map(function(n) { return n.trim(); }).filter(Boolean) : [];
        const result = readScheduleByYearMonth(sheetName, yearMonth, dateParam, names);
        if (!result.success) {
          return createJsonResponse({ success: false, error: result.error });
        }
        return createJsonResponse({
          success: true,
          records: result.records,
          details: result.details || {}
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
    // 解析請求內容
    const requestData = JSON.parse(e.postData.contents);
    const action = requestData.action;
    
    logDebug(`收到 POST 請求: ${action}`, { 
      action: action,
      timestamp: new Date().toISOString()
    });
    
    switch (action) {
      case 'upload':
        return handleUpload(requestData);
      
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
    if (uploadType === 'attendance') {
      return createJsonResponse({ success: false, error: '打卡上傳功能尚未實作' });
    }

    const fileName = requestData.fileName;
    const fileData = requestData.fileData;
    const targetSheetName = requestData.targetSheetName || '11501';
    const targetGoogleSheetTab = requestData.targetGoogleSheetTab || '班表';
    const branchName = requestData.branchName ? String(requestData.branchName).trim() : '';

    // 班表上傳時分店必選
    if (uploadType === 'schedule' && !branchName) {
      return createJsonResponse({ success: false, error: '請選擇分店' });
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
    
    // 準備轉換後的資料（parseResult.data 已經是解析好的格式，每列加 分店）
    logDebug('準備寫入資料');
    const headerRow = ['員工姓名', '排班日期', '上班時間', '下班時間', '工作時數', '班別', '分店'];
    const dataWithBranch = parseResult.data.map(function(row) {
      return row.concat([branchName]);
    });
    const transformedData = [headerRow].concat(dataWithBranch);
    logDebug(`共 ${parseResult.data.length} 列資料`);
    
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
      columns: ['員工姓名', '排班日期', '上班時間', '下班時間', '工作時數', '班別', '分店'],
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
        skippedCount: writeResult.skippedCount || 0
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
