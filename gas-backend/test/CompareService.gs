/**
 * 班表與打卡比對服務
 * 比對 key: 姓名 + 日期(YYYY-MM-DD) + 上班時間 + 下班時間 + 分店
 * 人員 sheet 對應：班表名稱 <-> 打卡名稱
 */

/**
 * 依分店取得人員名單（班表名稱、打卡名稱的聯集，供比對篩選用）
 * 人員 sheet: A=員工帳號, B=班表名稱, C=打卡名稱, D=分店
 * @param {string} branchName - 分店名稱
 * @return {Object} { success, names: [] }
 */
function getPersonnelByBranch(branchName) {
  try {
    if (!branchName || !branchName.toString().trim()) {
      return { success: true, names: [] };
    }
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.PERSONNEL || '人員';
    var allData = readFromSheet(sheetName);
    var nameSet = {};
    if (!allData || allData.length < 2) {
      return { success: true, names: [] };
    }
    var dataRows = allData.slice(1);
    var branchColIndex = 3;
    dataRows.forEach(function(row) {
      var b = row[branchColIndex] ? String(row[branchColIndex]).trim() : '';
      if (b !== branchName.toString().trim()) return;
      var attendanceName = row[2] ? String(row[2]).trim() : '';
      if (attendanceName) nameSet[attendanceName] = true;
    });
    var names = Object.keys(nameSet).sort();
    return { success: true, names: names };
  } catch (error) {
    logError('取得分店人員失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message, names: [] };
  }
}

/**
 * 讀取人員對應表（班表名稱、打卡名稱 <-> 員工帳號）
 * 人員 sheet: A=員工帳號, B=班表名稱, C=打卡名稱, D=分店
 * @return {Object} { scheduleNameToAccount: {}, attendanceNameToAccount: {}, accountToScheduleName: {}, accountToAttendanceName: {} }
 */
function readPersonnelMapping() {
  try {
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.PERSONNEL || '人員';
    var allData = readFromSheet(sheetName);
    var scheduleNameToAccount = {};
    var attendanceNameToAccount = {};
    var accountToScheduleName = {};
    var accountToAttendanceName = {};
    if (!allData || allData.length < 2) {
      return { scheduleNameToAccount: scheduleNameToAccount, attendanceNameToAccount: attendanceNameToAccount, accountToScheduleName: accountToScheduleName, accountToAttendanceName: accountToAttendanceName };
    }
    var dataRows = allData.slice(1);
    dataRows.forEach(function(row) {
      var acc = row[0] ? String(row[0]).trim() : '';
      var scheduleName = row[1] ? String(row[1]).trim() : '';
      var attendanceName = row[2] ? String(row[2]).trim() : '';
      if (acc) {
        if (scheduleName) { scheduleNameToAccount[scheduleName] = acc; accountToScheduleName[acc] = scheduleName; }
        if (attendanceName) { attendanceNameToAccount[attendanceName] = acc; accountToAttendanceName[acc] = attendanceName; }
      }
    });
    return { scheduleNameToAccount: scheduleNameToAccount, attendanceNameToAccount: attendanceNameToAccount, accountToScheduleName: accountToScheduleName, accountToAttendanceName: accountToAttendanceName };
  } catch (error) {
    logError('讀取人員對應失敗: ' + error.message, { error: error.toString() });
    return null;
  }
}

/**
 * 正規化日期為 YYYY-MM-DD
 */
function normalizeDateToDash(value) {
  if (!value) return '';
  var s = value.toString().trim();
  if (s.indexOf('/') >= 0) s = s.replace(/\//g, '-');
  if (s.length === 10 && s.indexOf('-') >= 0) return s;
  if (value instanceof Date && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var parsed = new Date(s);
  if (!isNaN(parsed)) {
    return Utilities.formatDate(parsed, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return s;
}

/**
 * 依條件讀取班表（支援月份或日期區間）
 * 班表: 員工姓名,排班日期,上班,下班,時數,班別,分店
 */
function readScheduleByConditions(yearMonth, startDate, endDate, names, branchName) {
  try {
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.SCHEDULE || '班表';
    var allData = readFromSheet(sheetName);
    if (!allData || allData.length < 2) {
      return { success: true, records: [] };
    }
    var dataRows = allData.slice(1);
    var dateColIndex = 1;
    var nameColIndex = 0;
    var branchColIndex = 6;
    var nameSet = (names && names.length > 0) ? {} : null;
    if (nameSet) names.forEach(function(n) { nameSet[String(n).trim()] = true; });
    var filtered = [];
    dataRows.forEach(function(row) {
      var dateVal = row[dateColIndex];
      if (!dateVal) return;
      var dateStr = normalizeDateToDash(dateVal);
      var matchDate = false;
      if (yearMonth && yearMonth.length === 6) {
        var prefix = yearMonth.substring(0, 4) + '-' + yearMonth.substring(4, 6);
        matchDate = dateStr.indexOf(prefix) === 0;
      } else if (startDate && startDate.length === 10) {
        var ed = endDate && endDate.length === 10 ? endDate : startDate;
        matchDate = dateStr >= startDate && dateStr <= ed;
      }
      if (!matchDate) return;
      if (branchName) {
        var b = row[branchColIndex] ? String(row[branchColIndex]).trim() : '';
        if (b !== branchName) return;
      }
      if (nameSet && Object.keys(nameSet).length > 0) {
        var n = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
        if (!nameSet[n]) return;
      }
      var r = row.slice();
      r[1] = dateStr;
      r[2] = normalizeTimeValue(row[2]);
      r[3] = normalizeTimeValue(row[3]);
      filtered.push(r);
    });
    return { success: true, records: filtered };
  } catch (error) {
    logError('讀取班表條件失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message, records: [] };
  }
}

/**
 * 依條件讀取打卡（支援月份或日期區間）
 * 打卡: 分店,員工編號,員工帳號,員工姓名,打卡日期,上班,下班,時數,狀態
 */
function readAttendanceByConditions(yearMonth, startDate, endDate, names, branchName) {
  try {
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.ATTENDANCE || '打卡';
    var allData = readFromSheet(sheetName);
    if (!allData || allData.length < 2) {
      return { success: true, records: [] };
    }
    var dataRows = allData.slice(1);
    var dateColIndex = 4;
    var nameColIndex = 3;
    var branchColIndex = 0;
    var nameSet = (names && names.length > 0) ? {} : null;
    if (nameSet) names.forEach(function(n) { nameSet[String(n).trim()] = true; });
    var filtered = [];
    dataRows.forEach(function(row) {
      var dateVal = row[dateColIndex];
      if (!dateVal) return;
      var dateStr = normalizeDateToDash(dateVal);
      var matchDate = false;
      if (yearMonth && yearMonth.length === 6) {
        var prefix = yearMonth.substring(0, 4) + '-' + yearMonth.substring(4, 6);
        matchDate = dateStr.indexOf(prefix) === 0;
      } else if (startDate && startDate.length === 10) {
        var ed = endDate && endDate.length === 10 ? endDate : startDate;
        matchDate = dateStr >= startDate && dateStr <= ed;
      }
      if (!matchDate) return;
      if (branchName) {
        var b = row[branchColIndex] ? String(row[branchColIndex]).trim() : '';
        if (b !== branchName) return;
      }
      if (nameSet && Object.keys(nameSet).length > 0) {
        var n = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
        if (!nameSet[n]) return;
      }
      var r = row.slice();
      r[4] = dateStr;
      r[5] = normalizeTimeValue(row[5]);
      r[6] = normalizeTimeValue(row[6]);
      filtered.push(r);
    });
    return { success: true, records: filtered };
  } catch (error) {
    logError('讀取打卡條件失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message, records: [] };
  }
}

/**
 * 建立比對 key: 員工帳號|日期|上班|下班|分店（用於校正紀錄）
 */
function buildCompareKey(empAccount, date, startTime, endTime, branch) {
  return [empAccount || '', date || '', startTime || '', endTime || '', branch || ''].join('|');
}

/**
 * 建立配對 key: 員工帳號|日期|分店（用於將班表與打卡配對為同一筆）
 * 同一人、同一日、同一分店只產生一筆，避免因上下班時間略異而拆成兩筆
 */
function buildMatchKey(empAccount, date, branch) {
  return [empAccount || '', date || '', branch || ''].join('|');
}

/**
 * 班表與打卡比對（一對一，無對應另一邊空白）
 * @return {Object} { success, items: [{ key, schedule, attendance, correction }] }
 */
function compareScheduleAttendance(yearMonth, startDate, endDate, names, branchName) {
  try {
    if (!branchName) {
      return { success: false, error: '請選擇分店' };
    }
    if ((!yearMonth || yearMonth.length !== 6) && (!startDate || startDate.length !== 10)) {
      return { success: false, error: '請選擇月份或日期區間' };
    }
    var mapping = readPersonnelMapping();
    if (!mapping) {
      return { success: false, error: '讀取人員對應失敗' };
    }
    var sResult = readScheduleByConditions(yearMonth, startDate, endDate, names, branchName);
    var aResult = readAttendanceByConditions(yearMonth, startDate, endDate, names, branchName);
    if (!sResult.success || !aResult.success) {
      return { success: false, error: sResult.error || aResult.error || '讀取失敗' };
    }
    var scheduleRecords = sResult.records || [];
    var attendanceRecords = aResult.records || [];
    var corrections = readCorrectionsValid(branchName);
    var correctionMap = {};
    if (corrections && corrections.length > 0) {
      corrections.forEach(function(c) {
        var k = buildCompareKey(c.empAccount, c.date, c.scheduleStart, c.scheduleEnd, c.branch);
        correctionMap[k] = c;
      });
    }
    var keyToSchedule = {};
    var keyToAttendance = {};
    var allKeys = {};
    scheduleRecords.forEach(function(row) {
      var sName = row[0] ? String(row[0]).trim() : '';
      var acc = mapping.attendanceNameToAccount[sName] || mapping.scheduleNameToAccount[sName] || '';
      var date = row[1] || '';
      var start = row[2] || '';
      var end = row[3] || '';
      var branch = row[6] ? String(row[6]).trim() : '';
      var key = buildMatchKey(acc || sName, date, branch);
      keyToSchedule[key] = { empAccount: acc, name: sName, date: date, startTime: start, endTime: end, hours: row[4], shift: row[5], branch: branch };
      allKeys[key] = true;
    });
    attendanceRecords.forEach(function(row) {
      var acc = row[2] ? String(row[2]).trim() : '';
      var aName = row[3] ? String(row[3]).trim() : '';
      var date = row[4] || '';
      var start = row[5] || '';
      var end = row[6] || '';
      var branch = row[0] ? String(row[0]).trim() : '';
      var key = buildMatchKey(acc || aName, date, branch);
      keyToAttendance[key] = { empAccount: acc, name: aName, date: date, startTime: start, endTime: end, hours: row[7], status: row[8], branch: branch };
      allKeys[key] = true;
    });
    var items = [];
    Object.keys(allKeys).forEach(function(key) {
      var s = keyToSchedule[key] || null;
      var a = keyToAttendance[key] || null;
      var empAcc = (s && s.empAccount) || (a && a.empAccount) || '';
      var date = (s && s.date) || (a && a.date) || '';
      var start = (s && s.startTime) || (a && a.startTime) || '';
      var end = (s && s.endTime) || (a && a.endTime) || '';
      var branch = (s && s.branch) || (a && a.branch) || branchName;
      var corrKey = buildCompareKey(empAcc, date, start, end, branch);
      var corr = correctionMap[corrKey] || null;
      var displayName = (a && a.name) ? a.name : (mapping.accountToAttendanceName[empAcc] || (s && s.name) || (a && a.name) || '');
      items.push({ key: key, schedule: s, attendance: a, correction: corr, displayName: displayName });
    });
    items.sort(function(x, y) {
      var d = (x.schedule && x.schedule.date) || (x.attendance && x.attendance.date) || '';
      var dy = (y.schedule && y.schedule.date) || (y.attendance && y.attendance.date) || '';
      if (d !== dy) return d.localeCompare(dy);
      var n = x.displayName || (x.schedule && x.schedule.name) || (x.attendance && x.attendance.name) || '';
      var ny = y.displayName || (y.schedule && y.schedule.name) || (y.attendance && y.attendance.name) || '';
      return n.localeCompare(ny);
    });
    return { success: true, items: items };
  } catch (error) {
    logError('班表打卡比對失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message };
  }
}

/**
 * 讀取有效校正紀錄（是否有效=是）
 * 校正 sheet 欄位：分店,員工帳號,姓名,日期,班表上班,班表下班,班表時數,打卡上班,打卡下班,打卡時數,打卡狀態,校正上班,校正下班,狀態,是否有效,建立時間
 */
function readCorrectionsValid(branchName) {
  try {
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.CORRECTION || '校正';
    var allData = readFromSheet(sheetName);
    if (!allData || allData.length < 2) return [];
    var header = allData[0];
    var validColIndex = header.indexOf('是否有效') >= 0 ? header.indexOf('是否有效') : 14;
    var dataRows = allData.slice(1);
    var result = [];
    dataRows.forEach(function(row) {
      var valid = row[validColIndex] ? String(row[validColIndex]).trim() : '';
      if (valid !== '是' && valid !== 'Y' && valid !== '1') return;
      var b = row[0] ? String(row[0]).trim() : '';
      if (branchName && b !== branchName) return;
      result.push({
        branch: b,
        empAccount: row[1] ? String(row[1]).trim() : '',
        name: row[2] ? String(row[2]).trim() : '',
        date: row[3] ? normalizeDateToDash(row[3]) : '',
        scheduleStart: row[4] ? String(row[4]).trim() : '',
        scheduleEnd: row[5] ? String(row[5]).trim() : '',
        correctedStart: row[11] ? String(row[11]).trim() : '',
        correctedEnd: row[12] ? String(row[12]).trim() : ''
      });
    });
    return result;
  } catch (error) {
    logError('讀取校正紀錄失敗: ' + error.message, { error: error.toString() });
    return [];
  }
}

/**
 * 寫入校正紀錄（若已有有效紀錄，設為無效後新增）
 */
function writeCorrection(data) {
  try {
    var config = getConfig();
    var sheetName = config.SHEET_NAMES.CORRECTION || '校正';
    var sheet = getOrCreateSheet(sheetName);
    var headerRow = ['分店', '員工帳號', '姓名', '日期', '班表上班', '班表下班', '班表時數', '打卡上班', '打卡下班', '打卡時數', '打卡狀態', '校正上班', '校正下班', '狀態', '是否有效', '建立時間'];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headerRow);
      formatHeaderRow(sheet, headerRow.length);
    }
    var validColIndex = 14;
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var existingData = sheet.getRange(2, 1, lastRow, Math.min(16, sheet.getLastColumn())).getValues();
      for (var idx = 0; idx < existingData.length; idx++) {
        var row = existingData[idx];
        var b = row[0] ? String(row[0]).trim() : '';
        var acc = row[1] ? String(row[1]).trim() : '';
        var date = row[3] ? normalizeDateToDash(row[3]) : '';
        var scheduleStart = row[4] ? String(row[4]).trim() : '';
        var scheduleEnd = row[5] ? String(row[5]).trim() : '';
        if (b === data.branch && acc === data.empAccount && date === data.date && scheduleStart === (data.scheduleStart || '') && scheduleEnd === (data.scheduleEnd || '')) {
          var validVal = row[validColIndex] ? String(row[validColIndex]).trim() : '';
          if (validVal === '是' || validVal === 'Y' || validVal === '1') {
            sheet.getRange(idx + 2, validColIndex + 1).setValue('否');
          }
        }
      }
    }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    var newRow = [
      data.branch || '',
      data.empAccount || '',
      data.name || '',
      data.date || '',
      data.scheduleStart || '',
      data.scheduleEnd || '',
      data.scheduleHours || '',
      data.attendanceStart || '',
      data.attendanceEnd || '',
      data.attendanceHours || '',
      data.attendanceStatus || '',
      data.correctedStart || '',
      data.correctedEnd || '',
      '校正',
      '是',
      now
    ];
    sheet.appendRow(newRow);
    logInfo('校正紀錄寫入成功', { branch: data.branch, empAccount: data.empAccount, date: data.date });
    return { success: true };
  } catch (error) {
    logError('寫入校正紀錄失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message };
  }
}
