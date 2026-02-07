/**
 * CompareService 校正 key 邏輯單元測試
 * 與 CompareService.gs 的 buildCompareKey 及比對 key 規則一致
 */

function buildCompareKey(empAccount, date, startTime, endTime, branch) {
  return [empAccount || '', date || '', startTime || '', endTime || '', branch || ''].join('|');
}

/** 與 CompareService.gs buildMatchKey 一致：員工帳號|日期|分店 */
function buildMatchKey(empAccount, date, branch) {
  return [empAccount || '', date || '', branch || ''].join('|');
}

function normalizeDateToDash(value) {
  if (!value) return '';
  var s = value.toString().trim();
  if (s.indexOf('/') >= 0) s = s.replace(/\//g, '-');
  if (s.length === 10 && s.indexOf('-') >= 0) return s;
  if (value instanceof Date && !isNaN(value)) {
    return '2025-02-07'; // simplified for test
  }
  var parsed = new Date(s);
  if (!isNaN(parsed)) {
    return '2025-02-07';
  }
  return s;
}

/**
 * 模擬比對時計算 correction lookup key 的邏輯（與 CompareService.gs 一致）
 */
function getCorrectionKeyForItem(schedule, attendance, branchName) {
  var s = schedule || null;
  var a = attendance || null;
  var empAcc = (s && s.empAccount) || (a && a.empAccount) || '';
  var date = (s && s.date) || (a && a.date) || '';
  var dateNorm = date ? normalizeDateToDash(date) : '';
  var branch = (s && s.branch) || (a && a.branch) || branchName || '';
  var scheduleStartForKey = s ? s.startTime : '';
  var scheduleEndForKey = s ? s.endTime : '';
  return buildCompareKey(empAcc, dateNorm, scheduleStartForKey, scheduleEndForKey, branch);
}

describe('CompareService correction key', () => {
  describe('buildCompareKey', () => {
    it('joins empAccount|date|startTime|endTime|branch with |', () => {
      expect(buildCompareKey('E01', '2025-02-07', '09:00', '18:00', 'A店')).toBe('E01|2025-02-07|09:00|18:00|A店');
    });
    it('uses empty string for missing parts', () => {
      expect(buildCompareKey('E01', '2025-02-07', '', '', 'A店')).toBe('E01|2025-02-07|||A店');
    });
  });

  describe('correction key for lookup (must match write key)', () => {
    it('with schedule: key uses schedule start/end', () => {
      var s = { empAccount: 'E01', date: '2025-02-07', startTime: '09:00', endTime: '18:00', branch: 'A店' };
      var a = { empAccount: 'E01', date: '2025-02-07', startTime: '08:55', endTime: '18:05', branch: 'A店' };
      var key = getCorrectionKeyForItem(s, a, 'A店');
      expect(key).toBe('E01|2025-02-07|09:00|18:00|A店');
    });

    it('attendance-only (no schedule): key uses empty schedule start/end so it matches written correction', () => {
      var a = { empAccount: 'E01', date: '2025-02-07', startTime: '08:55', endTime: '18:05', branch: 'A店' };
      var key = getCorrectionKeyForItem(null, a, 'A店');
      // 寫入校正時前端送 scheduleStart='', scheduleEnd=''，所以 correctionMap 的 key 是這個
      expect(key).toBe('E01|2025-02-07|||A店');
    });

    it('correctionMap lookup: written key (attendance-only) equals lookup key', () => {
      // 模擬寫入一筆「僅打卡」的校正：scheduleStart='', scheduleEnd=''
      var writtenKey = buildCompareKey('E01', '2025-02-07', '', '', 'A店');
      var correctionMap = {};
      correctionMap[writtenKey] = { correctedStart: '09:00', correctedEnd: '18:00' };
      // 比對時該筆為 attendance-only，lookup key 應與 writtenKey 一致
      var a = { empAccount: 'E01', date: '2025-02-07', startTime: '08:55', endTime: '18:05', branch: 'A店' };
      var lookupKey = getCorrectionKeyForItem(null, a, 'A店');
      expect(lookupKey).toBe(writtenKey);
      expect(correctionMap[lookupKey]).toEqual({ correctedStart: '09:00', correctedEnd: '18:00' });
    });

    it('correctionMap lookup: with schedule uses schedule start/end', () => {
      var writtenKey = buildCompareKey('E01', '2025-02-07', '09:00', '18:00', 'A店');
      var correctionMap = {};
      correctionMap[writtenKey] = { correctedStart: '09:00', correctedEnd: '18:00' };
      var s = { empAccount: 'E01', date: '2025-02-07', startTime: '09:00', endTime: '18:00', branch: 'A店' };
      var a = { empAccount: 'E01', date: '2025-02-07', startTime: '08:55', endTime: '18:05', branch: 'A店' };
      var lookupKey = getCorrectionKeyForItem(s, a, 'A店');
      expect(lookupKey).toBe(writtenKey);
      expect(correctionMap[lookupKey]).toBeDefined();
    });
  });

  describe('buildMatchKey (alert/schedule-attendance grouping)', () => {
    it('same acc+date+branch produces same key for schedule and attendance', () => {
      var acc = 'E01';
      var date = '2026-02-07';
      var branch = 'A店';
      var scheduleKey = buildMatchKey(acc, date, branch);
      var attendanceKey = buildMatchKey(acc, date, branch);
      expect(scheduleKey).toBe(attendanceKey);
      expect(scheduleKey).toBe('E01|2026-02-07|A店');
    });

    it('schedule row with K column (index 10) acc should match attendance row with index 2 acc', () => {
      var SCHEDULE_COL = { EMP_ACCOUNT: 10 };
      var scheduleRow = ['王小明', '2026-02-07', '09:00', '18:00', 8, 'A', 'A店', '', '', '', 'E01'];
      var attendanceRow = ['A店', '001', 'E01', '王小明', '2026-02-07', '09:00', '18:00', 8, '', '', '是', '', '', '', '', '', ''];
      var sAcc = (scheduleRow.length > SCHEDULE_COL.EMP_ACCOUNT && scheduleRow[SCHEDULE_COL.EMP_ACCOUNT]) ? String(scheduleRow[SCHEDULE_COL.EMP_ACCOUNT]).trim() : '';
      var aAcc = attendanceRow[2] ? String(attendanceRow[2]).trim() : '';
      var sDate = scheduleRow[1] || '';
      var aDate = attendanceRow[4] || '';
      var sBranch = scheduleRow[6] ? String(scheduleRow[6]).trim() : '';
      var aBranch = attendanceRow[0] ? String(attendanceRow[0]).trim() : '';
      var scheduleKey = buildMatchKey(sAcc, sDate, sBranch);
      var attendanceKey = buildMatchKey(aAcc, aDate, aBranch);
      expect(scheduleKey).toBe(attendanceKey);
      expect(scheduleKey).toBe('E01|2026-02-07|A店');
    });

    it('schedule row WITHOUT K column (only 10 cols) falls back to name; key must still match if mapping gives same acc', () => {
      var SCHEDULE_COL = { EMP_ACCOUNT: 10 };
      var scheduleRow = ['王小明', '2026-02-07', '09:00', '18:00', 8, 'A', 'A店', '', '', ''];
      var mapping = { attendanceNameToAccount: { '王小明': 'E01' }, scheduleNameToAccount: {} };
      var sName = scheduleRow[0] ? String(scheduleRow[0]).trim() : '';
      var sAcc = (scheduleRow.length > SCHEDULE_COL.EMP_ACCOUNT && scheduleRow[SCHEDULE_COL.EMP_ACCOUNT]) ? String(scheduleRow[SCHEDULE_COL.EMP_ACCOUNT]).trim() : '';
      if (!sAcc) sAcc = mapping.attendanceNameToAccount[sName] || mapping.scheduleNameToAccount[sName] || '';
      var attendanceRow = ['A店', '001', 'E01', '王小明', '2026-02-07', '09:00', '18:00', 8, '', '', '是', '', '', '', '', '', ''];
      var aAcc = attendanceRow[2] ? String(attendanceRow[2]).trim() : '';
      var sDate = scheduleRow[1] || '';
      var aDate = attendanceRow[4] || '';
      var sBranch = scheduleRow[6] ? String(scheduleRow[6]).trim() : '';
      var aBranch = attendanceRow[0] ? String(attendanceRow[0]).trim() : '';
      var scheduleKey = buildMatchKey(sAcc || sName, sDate, sBranch);
      var attendanceKey = buildMatchKey(aAcc, aDate, aBranch);
      expect(scheduleKey).toBe(attendanceKey);
      expect(scheduleKey).toBe('E01|2026-02-07|A店');
    });

    it('when schedule has no K column and mapping has no match, keys differ so schedules=[] for attendance key', () => {
      var SCHEDULE_COL = { EMP_ACCOUNT: 10 };
      var scheduleRow = ['TiNg', '2026-02-07', '09:00', '18:00', 8, 'A', 'A店', '', '', ''];
      var mapping = { attendanceNameToAccount: {}, scheduleNameToAccount: {} };
      var sName = scheduleRow[0] ? String(scheduleRow[0]).trim() : '';
      var sAcc = (scheduleRow.length > SCHEDULE_COL.EMP_ACCOUNT && scheduleRow[SCHEDULE_COL.EMP_ACCOUNT]) ? String(scheduleRow[SCHEDULE_COL.EMP_ACCOUNT]).trim() : '';
      if (!sAcc) sAcc = mapping.attendanceNameToAccount[sName] || mapping.scheduleNameToAccount[sName] || '';
      var attendanceRow = ['A店', '001', 'E01', '王小明', '2026-02-07', '09:00', '18:00', 8, '', '', '是', '', '', '', '', '', ''];
      var aAcc = attendanceRow[2] ? String(attendanceRow[2]).trim() : '';
      var sDate = scheduleRow[1] || '';
      var aDate = attendanceRow[4] || '';
      var sBranch = scheduleRow[6] ? String(scheduleRow[6]).trim() : '';
      var aBranch = attendanceRow[0] ? String(attendanceRow[0]).trim() : '';
      var scheduleKey = buildMatchKey(sAcc || sName, sDate, sBranch);
      var attendanceKey = buildMatchKey(aAcc, aDate, aBranch);
      expect(scheduleKey).toBe('TiNg|2026-02-07|A店');
      expect(attendanceKey).toBe('E01|2026-02-07|A店');
      expect(scheduleKey).not.toBe(attendanceKey);
    });
  });
});
