/**
 * 測試用：解析泉威國安班表
 * 使用 Node.js + xlsx 套件讀取 Excel 檔案
 */

const XLSX = require('xlsx');
const fs = require('fs');
const path = require('path');

// Excel 檔案路徑
const excelPath = 'c:\\Users\\FH4\\Downloads\\2026泉威國安班表.xlsx';

/**
 * 主函數：解析班表
 */
function runMasterShiftExtractor() {
  try {
    // 讀取 Excel 檔案
    const workbook = XLSX.readFile(excelPath);
    
    // 取得 11502 工作表
    const sheetName = '11502';
    if (!workbook.SheetNames.includes(sheetName)) {
      console.error(`找不到工作表 "${sheetName}"`);
      console.log('可用的工作表:', workbook.SheetNames);
      return;
    }
    
    const worksheet = workbook.Sheets[sheetName];
    
    // 轉換為二維陣列
    const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });
    
    console.log('=== Excel 檔案資訊 ===');
    console.log(`工作表名稱: ${sheetName}`);
    console.log(`總列數: ${data.length}`);
    console.log(`\n前 5 列資料預覽:`);
    data.slice(0, 5).forEach((row, idx) => {
      console.log(`第 ${idx + 1} 列:`, row.slice(0, 10));
    });
    
    // 1. 【動態建立代碼字典】
    console.log('\n=== 步驟 1: 掃描代碼定義 ===');
    const shiftMap = buildShiftMap(data);
    console.log('找到的班別代碼:', JSON.stringify(shiftMap, null, 2));
    
    // 2. 【定位資料區塊】
    console.log('\n=== 步驟 2: 定位資料結構 ===');
    const config = locateDataStructure(data);
    if (!config) {
      console.error('找不到日期列（1, 2, 3...），請確認工作表格式是否正確。');
      return;
    }
    console.log('資料結構配置:', config);
    console.log(`日期列位置: 第 ${config.dateRow + 1} 列, 第 ${config.dateCol + 1} 欄`);
    console.log(`員工姓名起始列: 第 ${config.startRow + 1} 列`);
    
    // 3. 【取得年月份】
    let yearMonth = data[1] && data[1][0] ? data[1][0].toString().trim() : '';
    console.log(`\n=== 步驟 3: 取得年月份 ===`);
    console.log(`原始年月: "${yearMonth}"`);
    if (!yearMonth.includes('/')) {
      yearMonth = '2026/01';
      console.log(`使用預設年月: ${yearMonth}`);
    }
    
    const finalResults = [];
    
    // 4. 【主解析迴圈】
    console.log(`\n=== 步驟 4: 解析班表資料 ===`);
    console.log(`開始從第 ${config.startRow + 1} 列解析...`);
    
    let processedEmployees = 0;
    
    for (let i = config.startRow; i < data.length; i++) {
      const empName = data[i] && data[i][0] ? data[i][0].toString().trim() : '';
      
      // 終止條件
      if (!empName || ['上班人數', '合計', '備註', 'P.T', '閉店評論'].some(k => empName.includes(k))) {
        if (i > config.startRow) break;
        continue;
      }
      
      processedEmployees++;
      console.log(`\n處理員工 ${processedEmployees}: ${empName} (第 ${i + 1} 列)`);
      
      let shiftCount = 0;
      
      // 遍歷所有日期欄位
      for (let j = config.dateCol; j < config.maxCol && j < (data[i] ? data[i].length : 0); j++) {
        const dayNum = data[config.dateRow] ? data[config.dateRow][j] : '';
        const cellValue = data[i][j] ? data[i][j].toString().trim() : '';
        
        if (!dayNum || isNaN(parseInt(dayNum)) || !cellValue || cellValue.toLowerCase() === 'nan') continue;
        
        let shiftInfo = { start: '', end: '', hours: '' };
        
        // A. 先查字典
        if (shiftMap[cellValue]) {
          shiftInfo = shiftMap[cellValue];
        }
        // B. 手寫時間
        else if (cellValue.includes('-')) {
          shiftInfo = parseTimeRange(cellValue);
        }
        else {
          continue;
        }
        
        if (shiftInfo.start) {
          const record = [
            empName,
            formatDate(yearMonth, dayNum),
            shiftInfo.start,
            shiftInfo.end,
            shiftInfo.hours
          ];
          finalResults.push(record);
          shiftCount++;
        }
      }
      
      console.log(`  → 找到 ${shiftCount} 個班次`);
    }
    
    // 5. 【輸出結果】
    console.log('\n=== 解析結果 ===');
    console.log(`總計提取 ${finalResults.length} 筆排班記錄`);
    console.log(`處理 ${processedEmployees} 位員工`);
    
    console.log('\n前 20 筆記錄:');
    console.log('員工姓名\t排班日期\t上班時間\t下班時間\t工作時數');
    console.log('─'.repeat(70));
    finalResults.slice(0, 20).forEach(record => {
      console.log(record.join('\t\t'));
    });
    
    if (finalResults.length > 20) {
      console.log(`\n... 還有 ${finalResults.length - 20} 筆記錄 ...`);
    }
    
    // 儲存為 JSON
    const outputPath = path.join(__dirname, 'parse-result.json');
    fs.writeFileSync(outputPath, JSON.stringify({
      shiftMap,
      config,
      totalRecords: finalResults.length,
      totalEmployees: processedEmployees,
      records: finalResults
    }, null, 2), 'utf-8');
    console.log(`\n完整結果已儲存至: ${outputPath}`);
    
    // 儲存為 CSV
    const csvPath = path.join(__dirname, 'parse-result.csv');
    const csvContent = [
      ['員工姓名', '排班日期', '上班時間', '下班時間', '工作時數'],
      ...finalResults
    ].map(row => row.join(',')).join('\n');
    fs.writeFileSync(csvPath, '\ufeff' + csvContent, 'utf-8'); // 加上 BOM 讓 Excel 正確顯示中文
    console.log(`CSV 結果已儲存至: ${csvPath}`);
    
  } catch (error) {
    console.error('解析失敗:', error.message);
    console.error(error.stack);
  }
}

/**
 * 輔助函數：定位資料結構
 */
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

/**
 * 輔助函數：建立代碼字典
 */
function buildShiftMap(data) {
  const map = {};
  data.forEach((row, rowIdx) => {
    const starIdx = row.indexOf('*');
    if (starIdx !== -1) {
      let code = '';
      let range = '';
      for (let i = starIdx + 1; i < row.length; i++) {
        let val = row[i] ? row[i].toString().trim() : '';
        if (val && val.length <= 3 && !code) code = val;
        if (val.includes('-') && /\d/.test(val)) range = val;
      }
      if (code && range) {
        map[code] = parseTimeRange(range);
        console.log(`  第 ${rowIdx + 1} 列: 代碼 "${code}" → ${range}`);
      }
    }
  });
  return map;
}

/**
 * 輔助函數：解析時間範圍
 */
function parseTimeRange(str) {
  try {
    const parts = str.split('-');
    const s = formatTime(parts[0]);
    const e = formatTime(parts[1]);
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

/**
 * 輔助函數：格式化時間
 */
function formatTime(t) {
  t = t.toString().replace(':', '').trim();
  return (t.length === 4) ? t.substring(0, 2) + ':' + t.substring(2, 4) : t;
}

/**
 * 輔助函數：格式化日期
 */
function formatDate(ym, d) {
  const dStr = (d < 10) ? '0' + d : d;
  return ym + '/' + dStr;
}

// 執行解析
runMasterShiftExtractor();
