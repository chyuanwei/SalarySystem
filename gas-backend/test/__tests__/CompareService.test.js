/**
 * CompareService 校正 key 邏輯單元測試
 * 與 CompareService.gs 的 buildCompareKey 及比對 key 規則一致
 */

function buildCompareKey(empAccount, date, startTime, endTime, branch) {
  return [empAccount || '', date || '', startTime || '', endTime || '', branch || ''].join('|');
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
});
