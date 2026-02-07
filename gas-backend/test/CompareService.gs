/**
 * 班表與打卡比對服務
 * 比對 key: 姓名 + 日期(YYYY-MM-DD) + 上班時間 + 下班時間 + 分店
 * 人員 sheet 對應：班表名稱 <-> 打卡名稱
 */

/**
 * 依該月份（或日期區間）＋分店的打卡資料取得人員名單。打卡 sheet 員工姓名欄直接取用。
 * @param {string} yearMonth - 年月 YYYYMM（與 startDate 二選一）
 * @param {string} startDate - 開始日期 YYYY-MM-DD
 * @param {string} endDate - 結束日期 YYYY-MM-DD
 * @param {string} branchName - 分店名稱（必填）
 * @return {Object} { success, names: [] }
 */
function getPersonnelFromAttendance(yearMonth, startDate, endDate, branchName) {
  try {
    if (!branchName || !branchName.toString().trim()) {
      return { success: true, names: [] };
    }
    if ((!yearMonth || yearMonth.length !== 6) && (!startDate || startDate.length !== 10)) {
      return { success: false, error: '請提供年月（YYYYMM）或日期區間', names: [] };
    }
    var aResult = readAttendanceByConditions(yearMonth, startDate, endDate, null, branchName);
    if (!aResult.success) {
      return { success: false, error: aResult.error || '讀取打卡失敗', names: [] };
    }
    var records = aResult.records || [];
    var nameSet = {};
    var nameColIndex = 3; // 打卡: 分店,員工編號,員工帳號,員工姓名,打卡日期,...
    records.forEach(function(row) {
      var n = row[nameColIndex] ? String(row[nameColIndex]).trim() : '';
      if (n) nameSet[n] = true;
    });
    var names = Object.keys(nameSet).sort();
    return { success: true, names: names };
  } catch (error) {
    logError('取得打卡人員失敗: ' + error.message, { error: error.toString() });
    return { success: false, error: error.message, names: [] };
  }
}

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
 * 將時間字串（HH:mm 或 HHmm）轉為當日分鐘數（0–1439），失敗回傳 null
 */
function timeStringToMinutes(t) {
  if (!t || typeof t !== 'string') return null;
  var s = String(t).trim().replace(':', '');
  if (!/^\d{3,4}$/.test(s)) return null;
  var h = parseInt(s.length === 4 ? s.substring(0, 2) : s.substring(0, 1), 10);
  var m = parseInt(s.length === 4 ? s.substring(2, 4) : s.substring(1, 4), 10);
  if (isNaN(h) || isNaN(m)) return null;
  return h * 60 + m;
}

/**
 * 從上班/下班時間計算區間分鐘數（跨日則 +24h）。失敗回傳 null。
 */
function timeRangeToMinutes(startStr, endStr) {
  var startM = timeStringToMinutes(startStr);
  var endM = timeStringToMinutes(endStr);
  if (startM === null || endM === null) return null;
  var diff = endM - startM;
  if (diff < 0) diff += 24 * 60;
  return diff;
}

/**
 * 時數欄位（數字或 "7.5"）轉為分鐘數，失敗回傳 null
 */
function hoursValueToMinutes(hours) {
  if (hours === undefined || hours === null || hours === '') return null;
  var n = parseFloat(String(hours).replace(/[^\d.]/g, ''));
  if (isNaN(n)) return null;
  return Math.round(n * 60);
}

/**
 * 兩時段是否重疊（依 start/end 分鐘數，跨日則 end+1440）
 */
function timeRangesOverlap(start1, end1, start2, end2) {
  var m1s = timeStringToMinutes(start1);
  var m1e = timeStringToMinutes(end1);
  var m2s = timeStringToMinutes(start2);
  var m2e = timeStringToMinutes(end2);
  if (m1s === null || m1e === null || m2s === null || m2e === null) return false;
  if (m1e < m1s) m1e += 24 * 60;
  if (m2e < m2s) m2e += 24 * 60;
  return (m1s < m2e && m2s < m1e);
}

/**
 * 依開始時間排序（用 timeStringToMinutes，無效排最後）
 */
function sortByStartTime(list) {
  list.sort(function(a, b) {
    var ma = timeStringToMinutes(a.startTime);
    var mb = timeStringToMinutes(b.startTime);
    if (ma === null) return 1;
    if (mb === null) return -1;
    return ma - mb;
  });
}

/**
 * 班表與打卡比對（同人同日同店可多筆；依開始時間 1-1 配對；時間重疊時設 overlapWarning）
 * @return {Object} { success, items: [{ key, schedule, attendance, correction, displayName, overtimeAlert?, overlapWarning? }] }
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
    var keyToSchedules = {};
    var keyToAttendances = {};
    var allKeys = {};
    scheduleRecords.forEach(function(row) {
      var sName = row[0] ? String(row[0]).trim() : '';
      var acc = mapping.attendanceNameToAccount[sName] || mapping.scheduleNameToAccount[sName] || '';
      var date = row[1] || '';
      var start = row[2] || '';
      var end = row[3] || '';
      var branch = row[6] ? String(row[6]).trim() : '';
      var key = buildMatchKey(acc || sName, date, branch);
      if (!keyToSchedules[key]) keyToSchedules[key] = [];
      keyToSchedules[key].push({ empAccount: acc, name: sName, date: date, startTime: start, endTime: end, hours: row[4], shift: row[5], branch: branch, remark: (row[7] !== undefined && row[7] !== null) ? String(row[7]).trim() : '' });
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
      if (!keyToAttendances[key]) keyToAttendances[key] = [];
      keyToAttendances[key].push({ empAccount: acc, name: aName, date: date, startTime: start, endTime: end, hours: row[7], status: row[8], branch: branch, remark: (row[9] !== undefined && row[9] !== null) ? String(row[9]).trim() : '' });
      allKeys[key] = true;
    });
    var config = getConfig();
    var overtimeAlertThreshold = (config.OVERTIME_ALERT !== undefined && config.OVERTIME_ALERT !== null) ? parseInt(config.OVERTIME_ALERT, 10) : 0;
    if (isNaN(overtimeAlertThreshold)) overtimeAlertThreshold = 0;
    var items = [];
    Object.keys(allKeys).forEach(function(key) {
      var schedules = keyToSchedules[key] || [];
      var attendances = keyToAttendances[key] || [];
      sortByStartTime(schedules);
      sortByStartTime(attendances);
      var scheduleOverlap = [];
      for (var i = 0; i < schedules.length; i++) {
        scheduleOverlap[i] = false;
        for (var j = 0; j < schedules.length; j++) {
          if (i !== j && timeRangesOverlap(schedules[i].startTime, schedules[i].endTime, schedules[j].startTime, schedules[j].endTime)) {
            scheduleOverlap[i] = true;
            break;
          }
        }
      }
      var attendanceOverlap = [];
      for (var ii = 0; ii < attendances.length; ii++) {
        attendanceOverlap[ii] = false;
        for (var jj = 0; jj < attendances.length; jj++) {
          if (ii !== jj && timeRangesOverlap(attendances[ii].startTime, attendances[ii].endTime, attendances[jj].startTime, attendances[jj].endTime)) {
            attendanceOverlap[ii] = true;
            break;
          }
        }
      }
      var maxLen = Math.max(schedules.length, attendances.length);
      for (var idx = 0; idx < maxLen; idx++) {
        var s = idx < schedules.length ? schedules[idx] : null;
        var a = idx < attendances.length ? attendances[idx] : null;
        var empAcc = (s && s.empAccount) || (a && a.empAccount) || '';
        var date = (s && s.date) || (a && a.date) || '';
        var dateNorm = date ? normalizeDateToDash(date) : '';
        var start = (s && s.startTime) || (a && a.startTime) || '';
        var end = (s && s.endTime) || (a && a.endTime) || '';
        var branch = (s && s.branch) || (a && a.branch) || branchName;
        // 校正 key 與寫入時一致：以班表上下班為 key（無班表時為空），否則僅打卡的校正會查不到
        var scheduleStartForKey = s ? s.startTime : '';
        var scheduleEndForKey = s ? s.endTime : '';
        var corrKey = buildCompareKey(empAcc, dateNorm, scheduleStartForKey, scheduleEndForKey, branch);
        var corr = correctionMap[corrKey] || null;
        var displayName = (a && a.name) ? a.name : (mapping.accountToAttendanceName[empAcc] || (s && s.name) || (a && a.name) || '');
        var overlapWarning = (s && idx < scheduleOverlap.length && scheduleOverlap[idx]) || (a && idx < attendanceOverlap.length && attendanceOverlap[idx]);
        var item = { key: key, schedule: s, attendance: a, correction: corr, displayName: displayName, overlapWarning: overlapWarning };
        if (s && a && overtimeAlertThreshold > 0) {
          var scheduleMins = timeRangeToMinutes(s.startTime, s.endTime);
          if (scheduleMins === null) scheduleMins = hoursValueToMinutes(s.hours);
          var attendanceMins = timeRangeToMinutes(a.startTime, a.endTime);
          if (attendanceMins === null) attendanceMins = hoursValueToMinutes(a.hours);
          if (scheduleMins !== null && attendanceMins !== null && (attendanceMins - scheduleMins) > overtimeAlertThreshold) {
            item.overtimeAlert = true;
          }
        }
        items.push(item);
      }
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
    var validColIndex = header.indexOf('是否有效') >= 0 ? header.indexOf('是否有效') : 15;
    var corrStartIdx = header.indexOf('校正上班') >= 0 ? header.indexOf('校正上班') : (header.indexOf('校正上班時間') >= 0 ? header.indexOf('校正上班時間') : 11);
    var corrEndIdx = header.indexOf('校正下班') >= 0 ? header.indexOf('校正下班') : (header.indexOf('校正下班時間') >= 0 ? header.indexOf('校正下班時間') : 12);
    var remarkIdx = header.indexOf('備註') >= 0 ? header.indexOf('備註') : 9;
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
        correctedStart: row[corrStartIdx] ? String(row[corrStartIdx]).trim() : '',
        correctedEnd: row[corrEndIdx] ? String(row[corrEndIdx]).trim() : '',
        remark: row[remarkIdx] ? String(row[remarkIdx]).trim() : ''
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
    var headerRow = ['分店', '員工帳號', '姓名', '日期', '班表上班', '班表下班', '班表時數', '打卡上班', '打卡下班', '打卡時數', '打卡狀態', '校正上班', '校正下班', '狀態', '備註', '是否有效', '建立時間'];
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(headerRow);
      formatHeaderRow(sheet, headerRow.length);
    }
    var validColIndex = 15;
    var lastRow = sheet.getLastRow();
    if (lastRow >= 2) {
      var existingData = sheet.getRange(2, 1, lastRow, Math.min(17, sheet.getLastColumn())).getValues();
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
      data.remark || data.correctionRemark || '',
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
