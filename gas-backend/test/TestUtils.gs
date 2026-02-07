/**
 * 測試工具函數
 * 用於測試新的泉威國安班表解析器
 */

/**
 * 測試解析泉威國安班表
 * 使用本地測試資料來驗證解析功能
 */
function testParseQuanWeiSchedule() {
  Logger.log('=== 開始測試泉威國安班表解析器 ===');
  
  // 模擬 11501 班表資料結構
  const testData = [
    ['', '更新日', '', '', '', '', '', '', '', ''],
    ['2026/01', '', '', '', '', '', '', '', '', ''],
    ['姓名/星期', '入職日', '上月剩', '', '', '', '', '', 1, 2, 3, 4, 5],
    ['', '', '餘年假', '週六', '週日', '週一', '週二', '週三', '週四', '週五', '週六', '週日', '週一'],
    ['TiNg', '', '', '', '', '', '', '', '', '', '', 'O', 'O'],
    ['茶葉', '', '', '', '', '', '', '', '', '', '', 'A', 'B'],
    ['魚', '', '', '', '', '', '', '', '', '', '', 'B1', 'A1'],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* A 10:00-17:00', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* A1 10:00-15:00', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* B 16:30-20:30', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* B1 14:30-20:30', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* O 10:00-20:30', '', '', '', '', '', '', '', '', '', '', '', '']
  ];
  
  try {
    // 測試 1: 定位資料結構
    Logger.log('\n--- 測試 1: 定位資料結構 ---');
    const config = locateDataStructure(testData);
    Logger.log('資料結構配置: ' + JSON.stringify(config));
    
    if (!config) {
      Logger.log('❌ 測試失敗：無法定位資料結構');
      return;
    }
    Logger.log('✅ 定位成功：日期列=' + (config.dateRow + 1) + ', 日期欄=' + (config.dateCol + 1));
    
    // 測試 2: 建立代碼字典
    Logger.log('\n--- 測試 2: 建立代碼字典 ---');
    const shiftMap = buildShiftMap(testData);
    Logger.log('代碼字典: ' + JSON.stringify(shiftMap, null, 2));
    
    const expectedCodes = ['A', 'A1', 'B', 'B1', 'O'];
    const foundCodes = Object.keys(shiftMap);
    if (expectedCodes.every(code => foundCodes.includes(code))) {
      Logger.log('✅ 代碼掃描成功：找到 ' + foundCodes.length + ' 個代碼');
    } else {
      Logger.log('❌ 代碼掃描不完整：預期 ' + expectedCodes.join(',') + '，實際 ' + foundCodes.join(','));
    }
    
    // 測試 3: 完整解析
    Logger.log('\n--- 測試 3: 完整解析 ---');
    const result = parseQuanWeiSchedule(testData, '測試工作表');
    Logger.log('解析結果:');
    Logger.log('  - 總記錄數: ' + result.records.length);
    Logger.log('  - 員工數: ' + result.totalEmployees);
    Logger.log('  - 班別代碼: ' + Object.keys(result.shiftMap).join(', '));
    
    Logger.log('\n前 5 筆記錄:');
    result.records.slice(0, 5).forEach((record, idx) => {
      Logger.log('  ' + (idx + 1) + '. ' + record.join(' | '));
    });
    
    if (result.records.length > 0) {
      Logger.log('✅ 解析成功');
    } else {
      Logger.log('❌ 解析失敗：沒有產生記錄');
    }
    
    // 測試 4: 時間解析
    Logger.log('\n--- 測試 4: 時間解析 ---');
    const testTimeRanges = [
      '10:00-17:00',
      '16:30-20:30',
      '22:00-06:00' // 跨日
    ];
    
    testTimeRanges.forEach(range => {
      const parsed = parseTimeRange(range);
      Logger.log(range + ' → ' + JSON.stringify(parsed));
    });
    Logger.log('✅ 時間解析測試完成');
    
    // 測試 5: 日期格式化
    Logger.log('\n--- 測試 5: 日期格式化 ---');
    const testDates = [
      { ym: '2026/01', d: 1, expected: '2026-01-01' },
      { ym: '2026/02', d: 15, expected: '2026-02-15' },
      { ym: '2026/12', d: 31, expected: '2026-12-31' },
      { ym: '202601', d: 7, expected: '2026-01-07' }
    ];
    
    let dateTestPass = true;
    testDates.forEach(test => {
      const formatted = formatScheduleDate(test.ym, test.d);
      const pass = (formatted === test.expected);
      Logger.log((pass ? '✅' : '❌') + ' ' + test.ym + ' + ' + test.d + ' = ' + formatted + ' (預期: ' + test.expected + ')');
      if (!pass) dateTestPass = false;
    });
    
    Logger.log('\n=== 測試完成 ===');
    Logger.log('總體結果: ' + (dateTestPass ? '✅ 全部通過' : '⚠️ 部分測試失敗'));
    
  } catch (error) {
    Logger.log('❌ 測試過程發生錯誤: ' + error.message);
    Logger.log('錯誤堆疊: ' + error.stack);
  }
}

/**
 * 測試完整的上傳流程（模擬）
 */
function testFullUploadFlow() {
  Logger.log('=== 測試完整上傳流程 ===');
  
  try {
    // 這個測試需要實際的 Excel 檔案
    // 由於無法在測試環境中直接讀取檔案，這裡只記錄測試步驟
    
    Logger.log('測試步驟:');
    Logger.log('1. 準備測試 Excel 檔案（2026泉威國安班表.xlsx）');
    Logger.log('2. 使用前端上傳介面上傳檔案');
    Logger.log('3. 選擇工作表 11501 或 11502');
    Logger.log('4. 檢查 Google Sheets 中的「國安班表」工作表');
    Logger.log('5. 檢查 Log 工作表是否有正確的記錄');
    Logger.log('6. 驗證資料格式：員工姓名 | 排班日期 | 上班時間 | 下班時間 | 工作時數');
    
    Logger.log('\n請手動執行上述步驟進行完整測試');
    
  } catch (error) {
    Logger.log('❌ 測試失敗: ' + error.message);
  }
}

/**
 * 檢查 GAS 環境設定
 */
function checkEnvironmentSetup() {
  Logger.log('=== 檢查環境設定 ===');
  
  try {
    // 檢查 Config
    const config = getConfig();
    Logger.log('✅ Config 讀取成功');
    Logger.log('  SHEET_ID: ' + config.SHEET_ID);
    Logger.log('  ENVIRONMENT: ' + config.ENVIRONMENT);
    
    // 檢查 Google Sheets 連線
    const ss = getSpreadsheet();
    Logger.log('✅ Google Sheets 連線成功');
    Logger.log('  試算表名稱: ' + ss.getName());
    Logger.log('  試算表 ID: ' + ss.getId());
    
    // 檢查工作表
    const sheets = ss.getSheets();
    Logger.log('✅ 工作表列表:');
    sheets.forEach((sheet, idx) => {
      Logger.log('  ' + (idx + 1) + '. ' + sheet.getName());
    });
    
    // 檢查 Log 等級
    const logLevel = getLogLevel();
    Logger.log('✅ Log 等級: ' + logLevel + ' (' + 
      (logLevel === 0 ? 'OFF' : logLevel === 1 ? 'OPERATION' : 'DEBUG') + ')');
    
    Logger.log('\n=== 環境檢查完成 ===');
    
  } catch (error) {
    Logger.log('❌ 環境檢查失敗: ' + error.message);
    Logger.log('請檢查 Script Properties 是否正確設定 SHEET_ID');
  }
}
