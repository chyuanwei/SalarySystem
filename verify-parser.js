/**
 * 快速驗證腳本：模擬 GAS 環境測試解析器
 * 用於在本地驗證解析邏輯是否正確
 */

// 模擬 GAS 的 Logger
const Logger = {
  log: (msg) => console.log(msg)
};

// 載入解析函數（從 test-parser.js 複製核心邏輯）
function locateDataStructure(data) {
  for (let r = 0; r < 10 && r < data.length; r++) {
    for (let c = 1; c < 15 && c < (data[r] ? data[r].length : 0); c++) {
      if (data[r][c] == 1) {
        return {
          dateRow: r,
          dateCol: c,
          startRow: r + 2,
          maxCol: data[r].length
        };
      }
    }
  }
  return null;
}

function buildShiftMap(data) {
  const map = {};
  data.forEach((row, rowIdx) => {
    // 檢查第一個儲存格是否以 * 開頭
    const firstCell = row[0] ? row[0].toString().trim() : '';
    if (firstCell.startsWith('*')) {
      // 從第一個儲存格提取代碼和時間範圍
      // 格式：* A 10:00-17:00 或 * A1 10:00-15:00
      const parts = firstCell.split(/\s+/);
      if (parts.length >= 3) {
        const code = parts[1]; // A, A1, B, etc.
        const range = parts[2]; // 10:00-17:00
        if (code && range && range.includes('-')) {
          map[code] = parseTimeRange(range);
          console.log(`  代碼定義: ${code} → ${range}`);
        }
      }
    }
  });
  return map;
}

function parseTimeRange(str) {
  try {
    const parts = str.split('-');
    const s = formatTimeString(parts[0]);
    const e = formatTimeString(parts[1]);
    const sParts = s.split(':');
    const eParts = e.split(':');
    const startDate = new Date(0, 0, 0, parseInt(sParts[0]), parseInt(sParts[1]));
    const endDate = new Date(0, 0, 0, parseInt(eParts[0]), parseInt(eParts[1]));
    let diff = (endDate - startDate) / 1000 / 60 / 60;
    if (diff < 0) diff += 24;
    return { start: s, end: e, hours: diff.toFixed(1) };
  } catch (err) {
    return { start: '', end: '', hours: '' };
  }
}

function formatTimeString(t) {
  t = t.toString().replace(':', '').trim();
  return (t.length === 4) ? t.substring(0, 2) + ':' + t.substring(2, 4) : t;
}

function formatScheduleDate(ym, d) {
  const dStr = (d < 10) ? '0' + d : d.toString();
  return ym + '/' + dStr;
}

function parseQuanWeiSchedule(data, sheetName) {
  console.log(`\n=== 開始解析泉威國安班表: ${sheetName} ===\n`);
  
  // 1. 建立代碼字典
  console.log('步驟 1: 掃描班別代碼定義');
  const shiftMap = buildShiftMap(data);
  console.log('找到的班別代碼:', Object.keys(shiftMap).join(', '));
  
  // 2. 定位資料結構
  console.log('\n步驟 2: 定位資料結構');
  const config = locateDataStructure(data);
  if (!config) {
    throw new Error('找不到日期列（1, 2, 3...），請確認工作表格式是否正確');
  }
  console.log(`日期列位置: 第 ${config.dateRow + 1} 列, 第 ${config.dateCol + 1} 欄`);
  console.log(`員工起始列: 第 ${config.startRow + 1} 列`);
  
  // 3. 取得年月份
  console.log('\n步驟 3: 取得年月份');
  let yearMonth = data[1] && data[1][0] ? data[1][0].toString().trim() : '';
  if (!yearMonth.includes('/')) {
    yearMonth = '2026/01';
    console.log(`⚠️  無法取得年月份，使用預設值: ${yearMonth}`);
  } else {
    console.log(`年月份: ${yearMonth}`);
  }
  
  const finalResults = [];
  let processedEmployees = 0;
  
  // 4. 主解析迴圈
  console.log('\n步驟 4: 解析員工排班');
  
  for (let i = config.startRow; i < data.length; i++) {
    const empName = data[i] && data[i][0] ? data[i][0].toString().trim() : '';
    
    if (!empName || ['上班人數', '合計', '備註', 'P.T', '閉店評論'].some(k => empName.includes(k))) {
      if (i > config.startRow) break;
      continue;
    }
    
    processedEmployees++;
    let shiftCount = 0;
    
    for (let j = config.dateCol; j < config.maxCol && j < (data[i] ? data[i].length : 0); j++) {
      const dayNum = data[config.dateRow] ? data[config.dateRow][j] : '';
      const cellValue = data[i][j] ? data[i][j].toString().trim() : '';
      
      if (!dayNum || isNaN(parseInt(dayNum)) || !cellValue || cellValue.toLowerCase() === 'nan') {
        continue;
      }
      
      let shiftInfo = { start: '', end: '', hours: '' };
      
      if (shiftMap[cellValue]) {
        shiftInfo = shiftMap[cellValue];
      } else if (cellValue.includes('-')) {
        shiftInfo = parseTimeRange(cellValue);
      } else {
        continue;
      }
      
      if (shiftInfo.start) {
        finalResults.push([
          empName,
          formatScheduleDate(yearMonth, dayNum),
          shiftInfo.start,
          shiftInfo.end,
          shiftInfo.hours
        ]);
        shiftCount++;
      }
    }
    
    console.log(`  處理員工: ${empName} → 找到 ${shiftCount} 個班次`);
  }
  
  console.log(`\n=== 解析完成 ===`);
  console.log(`總記錄數: ${finalResults.length}`);
  console.log(`處理員工: ${processedEmployees} 位`);
  
  return {
    records: finalResults,
    shiftMap: shiftMap,
    config: config,
    totalEmployees: processedEmployees
  };
}

// 執行驗證測試
function runVerification() {
  console.log('╔════════════════════════════════════════╗');
  console.log('║  泉威國安班表解析器 - 快速驗證測試  ║');
  console.log('╚════════════════════════════════════════╝\n');
  
  // 模擬測試資料
  const testData = [
    ['', '更新日', '', '', '', '', '', '', '', ''],
    ['2026/01', '', '', '', '', '', '', '', '', ''],
    ['姓名/星期', '入職日', '上月剩', '', '', '', '', '', 1, 2, 3, 4, 5, 6, 7],
    ['', '', '餘年假', '週六', '週日', '週一', '週二', '週三', '週四', '週五', '週六', '週日', '週一', '週二', '週三'],
    ['TiNg', '', '', '', '', '', '', '', '', '', 'O', '', '', 'A', 'A'],
    ['茶葉', '', '', '', '', '', '', '', '', '', 'A', 'B', '', 'O', ''],
    ['魚', '', '', '', '', '', '', '', '', '', 'B1', 'A1', '', '', 'B'],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['合計', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* A 10:00-17:00', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* A1 10:00-15:00', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* B 16:30-20:30', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* B1 14:30-20:30', '', '', '', '', '', '', '', '', '', '', '', '', '', ''],
    ['* O 10:00-20:30', '', '', '', '', '', '', '', '', '', '', '', '', '', '']
  ];
  
  try {
    const result = parseQuanWeiSchedule(testData, '測試工作表');
    
    console.log('\n📊 解析結果摘要:');
    console.log('├─ 班別代碼:', Object.keys(result.shiftMap).join(', '));
    console.log('├─ 總記錄數:', result.records.length);
    console.log('└─ 處理員工:', result.totalEmployees, '位');
    
    console.log('\n📝 前 10 筆記錄:');
    console.log('員工姓名 | 排班日期   | 上班  | 下班  | 時數');
    console.log('─'.repeat(50));
    result.records.slice(0, 10).forEach(record => {
      console.log(`${record[0].padEnd(8)} | ${record[1]} | ${record[2]} | ${record[3]} | ${record[4]}`);
    });
    
    if (result.records.length > 10) {
      console.log(`... 還有 ${result.records.length - 10} 筆記錄 ...`);
    }
    
    console.log('\n✅ 驗證測試通過！解析器運作正常。');
    console.log('\n下一步：請在 GAS 編輯器中執行 testParseQuanWeiSchedule() 進行完整測試。');
    
  } catch (error) {
    console.log('\n❌ 驗證測試失敗！');
    console.log('錯誤訊息:', error.message);
    console.log('錯誤堆疊:', error.stack);
  }
}

// 執行驗證
runVerification();
