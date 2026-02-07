/**
 * 薪資計算系統 - 打卡 CSV 解析器
 * 功能：解析打卡 CSV 格式（員工編號、員工姓名、打卡日期、上班/下班時間、打卡地點等）
 */

/**
 * 解析打卡 CSV 檔案
 * CSV 欄位：員工編號、員工姓名、員工帳號、打卡日期、上班/下班時間、打卡地點、工作時數、狀態
 * @param {string} base64Data - Base64 編碼的 CSV 檔案內容
 * @param {string} fileName - 檔案名稱
 * @return {Object} { success, records: [{ 員工編號, 員工姓名, 員工帳號, 打卡日期, 上班時間, 下班時間, 打卡地點, 工作時數, 狀態 }], error }
 */
function parseAttendanceCSV(base64Data, fileName) {
  try {
    logDebug('開始解析打卡 CSV: ' + fileName, { fileName: fileName });
    
    var decodedBytes = Utilities.base64Decode(base64Data);
    var csvContent = Utilities.newBlob(decodedBytes).getDataAsString('UTF-8');
    
    var lines = csvContent.split(/\r?\n/).filter(function(line) { return line.trim().length > 0; });
    if (lines.length < 2) {
      return { success: false, error: 'CSV 檔案無資料列' };
    }
    
    var headerLine = lines[0];
    var dataLines = lines.slice(1);
    var records = [];
    
    for (var i = 0; i < dataLines.length; i++) {
      var fields = parseCSVLine(dataLines[i]);
      if (fields.length < 6) continue;
      
      var empNo = (fields[0] || '').toString().trim();
      var empName = (fields[1] || '').toString().trim();
      var empAccount = stripExcelFormula((fields[2] || '').toString().trim());
      var punchDate = (fields[3] || '').toString().trim();
      var timeRange = (fields[4] || '').toString().trim();
      var location = (fields[5] || '').toString().trim();
      var workHours = (fields[6] !== undefined) ? (fields[6] || '').toString().trim() : '';
      var status = (fields[7] !== undefined) ? (fields[7] || '').toString().trim() : '';
      
      var startTime = '';
      var endTime = '';
      if (timeRange.indexOf('~') >= 0) {
        var parts = timeRange.split('~');
        startTime = (parts[0] || '').trim();
        endTime = (parts[1] || '').trim();
      }
      
      records.push({
        empNo: empNo,
        empName: empName,
        empAccount: empAccount,
        punchDate: punchDate,
        startTime: startTime,
        endTime: endTime,
        location: location,
        workHours: workHours,
        status: status
      });
    }
    
    logInfo('打卡 CSV 解析成功', { fileName: fileName, recordCount: records.length });
    return { success: true, records: records };
    
  } catch (error) {
    logError('打卡 CSV 解析失敗: ' + error.message, { fileName: fileName, error: error.toString() });
    return { success: false, error: 'CSV 解析失敗: ' + error.message };
  }
}

/**
 * 解析單行 CSV（處理雙引號內的逗號）
 */
function parseCSVLine(line) {
  var result = [];
  var current = '';
  var inQuotes = false;
  
  for (var i = 0; i < line.length; i++) {
    var c = line.charAt(i);
    if (c === '"') {
      inQuotes = !inQuotes;
    } else if (c === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += c;
    }
  }
  result.push(current);
  return result;
}

/**
 *  stripping Excel 公式格式（如 =""669929"" -> 669929）
 */
function stripExcelFormula(value) {
  if (!value) return '';
  var s = value.trim();
  if (s.indexOf('=') === 0) {
    s = s.substring(1).replace(/^["']|["']$/g, '');
  }
  return s.replace(/^["']|["']$/g, '').trim();
}
