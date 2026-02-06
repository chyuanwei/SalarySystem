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
    
    logToSheet(`收到 POST 請求: ${action}`, 'INFO');
    
    switch (action) {
      case 'upload':
        return handleUpload(requestData);
      
      case 'calculate':
        return handleCalculate(requestData);
      
      default:
        return createJsonResponse({ 
          success: false, 
          error: `未知的 action: ${action}` 
        });
    }
    
  } catch (error) {
    logToSheet(`POST 請求處理失敗: ${error.message}`, 'ERROR');
    return createErrorResponse(error);
  }
}

/**
 * 處理檔案上傳
 */
function handleUpload(requestData) {
  try {
    const fileName = requestData.fileName;
    const fileData = requestData.fileData;
    const targetSheetName = requestData.targetSheetName || '11501';
    const targetGoogleSheetTab = requestData.targetGoogleSheetTab || '國安班表';
    
    logToSheet(`開始處理上傳: ${fileName}, 目標工作表: ${targetSheetName}`, 'INFO');
    
    // 驗證必要參數
    if (!fileName || !fileData) {
      throw new Error('缺少必要參數: fileName 或 fileData');
    }
    
    // 解析 Excel 檔案
    const parseResult = parseExcel(fileData, fileName, targetSheetName);
    
    if (!parseResult.success) {
      throw new Error(parseResult.error);
    }
    
    // 驗證資料
    const validation = validateExcelData(parseResult.data);
    
    if (!validation.valid) {
      throw new Error('資料驗證失敗: ' + validation.errors.join(', '));
    }
    
    // 記錄警告
    if (validation.warnings && validation.warnings.length > 0) {
      logToSheet(`資料警告: ${validation.warnings.join('; ')}`, 'WARNING');
    }
    
    // 寫入 Google Sheets
    const writeResult = writeToSheet(parseResult.data, targetGoogleSheetTab);
    
    if (!writeResult.success) {
      throw new Error(writeResult.error);
    }
    
    // 建立處理記錄
    createProcessRecord({
      fileName: fileName,
      sourceSheet: targetSheetName,
      targetSheet: targetGoogleSheetTab,
      rowCount: parseResult.rowCount,
      status: 'SUCCESS',
      message: '檔案上傳並處理成功'
    });
    
    // 回傳成功訊息
    return createJsonResponse({
      success: true,
      message: '檔案已成功上傳並處理',
      details: {
        fileName: fileName,
        sourceSheet: targetSheetName,
        targetSheet: targetGoogleSheetTab,
        rowCount: parseResult.rowCount,
        columnCount: writeResult.columnCount,
        warnings: validation.warnings || []
      }
    });
    
  } catch (error) {
    const errorMsg = `檔案上傳處理失敗: ${error.message}`;
    logToSheet(errorMsg, 'ERROR');
    
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
