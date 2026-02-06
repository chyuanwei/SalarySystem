# Log 系統流程圖

## 📊 完整 Log 記錄流程

```
┌─────────────────────────────────────────────────────────────────┐
│                     應用程式呼叫 Log 函數                        │
│                                                                 │
│  logDebug()  logInfo()  logWarning()  logError()  logOperation()│
└───────────────────────────┬─────────────────────────────────────┘
                           │
                           ▼
         ┌─────────────────────────────────────┐
         │      logToSheet(message, level,     │
         │             details)                │
         └─────────────────┬───────────────────┘
                          │
                          ▼
         ┌─────────────────────────────────────┐
         │   步驟 1: 取得 Log 等級設定          │
         │   getLogLevel()                     │
         │   從 Script Properties 讀取         │
         │   Log_Level (0/1/2)                │
         └─────────────────┬───────────────────┘
                          │
                          ▼
         ┌─────────────────────────────────────┐
         │   步驟 2: 檢查是否關閉               │
         │   if (Log_Level === 0)              │
         └─────────────────┬───────────────────┘
                          │
                ┌─────────┴─────────┐
                │                   │
               Yes                 No
                │                   │
         ┌──────▼──────┐           │
         │  直接返回    │           │
         │  不記錄      │           │
         └─────────────┘           │
                                   ▼
                    ┌──────────────────────────────┐
                    │   步驟 3: 判斷是否記錄        │
                    │   checkShouldLog(level,      │
                    │                  current)    │
                    └──────────────┬───────────────┘
                                   │
                    ┌──────────────┴──────────────┐
                    │                             │
               Level 1 (OPERATION)           Level 2 (DEBUG)
                    │                             │
         ┌──────────▼──────────┐      ┌──────────▼──────────┐
         │  只記錄:              │      │  記錄所有等級:       │
         │  - OPERATION         │      │  - DEBUG            │
         │  - ERROR             │      │  - INFO             │
         │                      │      │  - WARNING          │
         │  跳過:                │      │  - ERROR            │
         │  - DEBUG             │      │  - OPERATION        │
         │  - INFO              │      └──────────┬──────────┘
         │  - WARNING           │                 │
         └──────────┬───────────┘                 │
                    │                             │
                    └──────────┬──────────────────┘
                               │
                               ▼
                    ┌─────────────────────────────┐
                    │   步驟 4: 取得 Log 工作表    │
                    │   getOrCreateSheet('Log')   │
                    └─────────────┬───────────────┘
                                  │
                                  ▼
                    ┌─────────────────────────────┐
                    │   步驟 5: 檢查是否初始化     │
                    │   if (lastRow === 0)        │
                    └─────────────┬───────────────┘
                                  │
                        ┌─────────┴─────────┐
                        │                   │
                       Yes                 No
                        │                   │
         ┌──────────────▼──────────┐       │
         │   initLogSheet()         │       │
         │   - 建立標題列           │       │
         │   - 格式化（深灰底白字） │       │
         │   - 凍結標題列           │       │
         │   - 設定欄寬             │       │
         └──────────────┬──────────┘       │
                        │                   │
                        └─────────┬─────────┘
                                  │
                                  ▼
                    ┌─────────────────────────────┐
                    │   步驟 6: 建立 Log 記錄      │
                    │   logEntry = [              │
                    │     timestamp,              │
                    │     level,                  │
                    │     message,                │
                    │     environment,            │
                    │     user,                   │
                    │     JSON.stringify(details) │
                    │   ]                         │
                    └─────────────┬───────────────┘
                                  │
                                  ▼
                    ┌─────────────────────────────┐
                    │   步驟 7: 寫入 Sheets        │
                    │   sheet.appendRow(logEntry) │
                    └─────────────┬───────────────┘
                                  │
                                  ▼
                    ┌─────────────────────────────┐
                    │   步驟 8: 輸出到 Logger      │
                    │   Logger.log(`[${level}]    │
                    │              ${message}`)   │
                    │                             │
                    │   if (details)              │
                    │     Logger.log(details)     │
                    └─────────────────────────────┘
```

---

## 🔄 Log 等級過濾決策樹

```
呼叫 logToSheet(message, level, details)
           │
           ▼
    Log_Level = ?
           │
    ┌──────┼──────┐
    │      │      │
   0=OFF  1=OP   2=DEBUG
    │      │      │
    ▼      │      │
 ┌─────┐  │      │
 │不記錄│  │      │
 └─────┘  │      │
          ▼      ▼
       level=? 記錄所有
          │
    ┌─────┼─────┐
    │     │     │
  ERROR  OP   其他
    │     │     │
    ▼     ▼     ▼
  記錄   記錄  不記錄
```

---

## 📝 實際使用案例

### 案例 1: 檔案上傳流程

```javascript
function handleUpload(requestData) {
  const startTime = new Date();
  
  // 1. 開始操作記錄
  logOperation(`開始處理上傳: ${fileName}`, {
    fileName: fileName,
    targetSheetName: targetSheetName
  });
  
  // 2. 除錯資訊
  logDebug(`檔案大小: ${fileData.length} bytes (Base64)`);
  
  // 3. 解析成功
  const parseResult = parseExcel(...);
  logDebug(`Excel 解析成功，讀取 ${parseResult.rowCount} 列資料`);
  
  // 4. 驗證警告
  if (validation.warnings.length > 0) {
    logWarning(`資料警告`, { warnings: validation.warnings });
  }
  
  // 5. 處理完成
  const processTime = (new Date() - startTime) / 1000;
  logOperation(`檔案處理成功: ${fileName}`, {
    rowCount: rowCount,
    processTime: `${processTime.toFixed(2)} 秒`
  });
}
```

**產生的 Log 記錄（Log_Level = 2）：**
```
2026-02-06 10:30:00 | OPERATION | 開始處理上傳: test.xlsx | test | user@... | {"fileName":"test.xlsx"}
2026-02-06 10:30:01 | DEBUG     | 檔案大小: 12345 bytes   | test | user@... | {}
2026-02-06 10:30:03 | DEBUG     | Excel 解析成功          | test | user@... | {"rowCount":21}
2026-02-06 10:30:05 | OPERATION | 檔案處理成功            | test | user@... | {"rowCount":21,"processTime":"5.2 秒"}
```

**產生的 Log 記錄（Log_Level = 1）：**
```
2026-02-06 10:30:00 | OPERATION | 開始處理上傳: test.xlsx | test | user@... | {"fileName":"test.xlsx"}
2026-02-06 10:30:05 | OPERATION | 檔案處理成功            | test | user@... | {"rowCount":21,"processTime":"5.2 秒"}
```
（DEBUG 等級被過濾掉）

---

### 案例 2: 泉威班表解析流程

```javascript
function parseQuanWeiSchedule(data, sheetName) {
  // 1. 重要操作開始
  logOperation(`開始解析泉威國安班表: ${sheetName}`);
  
  // 2. 代碼掃描結果
  const shiftMap = buildShiftMap(data);
  logDebug('找到的班別代碼', { 
    shiftMap: shiftMap,
    codeCount: Object.keys(shiftMap).length
  });
  
  // 3. 結構定位
  const config = locateDataStructure(data);
  logDebug('資料結構配置', { config: config });
  
  // 4. 警告處理
  if (!yearMonth.includes('/')) {
    logWarning('無法取得年月份，使用預設值', { 
      yearMonth: yearMonth 
    });
  }
  
  // 5. 解析過程
  for (每個員工) {
    logDebug(`處理員工: ${empName}`, { shiftCount: shiftCount });
  }
  
  // 6. 完成操作
  logOperation(`泉威國安班表解析完成`, {
    sheetName: sheetName,
    totalRecords: finalResults.length,
    totalEmployees: processedEmployees,
    shiftCodes: Object.keys(shiftMap).join(', ')
  });
}
```

**產生的完整 Log 記錄（Log_Level = 2）：**
```
[OPERATION] 開始解析泉威國安班表: 11501
  Details: {}
  
[DEBUG] 找到的班別代碼
  Details: {"shiftMap":{"A":{"start":"10:00","end":"17:00","hours":"7.0"},...},"codeCount":5}
  
[DEBUG] 資料結構配置
  Details: {"config":{"dateRow":2,"dateCol":8,"startRow":4,"maxCol":260}}
  
[DEBUG] 年月份
  Details: {"yearMonth":"2026/01"}
  
[DEBUG] 開始從第 5 列解析員工排班
  Details: {}
  
[DEBUG] 處理員工: TiNg
  Details: {"shiftCount":4}
  
[DEBUG] 處理員工: 茶葉
  Details: {"shiftCount":9}
  
[DEBUG] 處理員工: 魚
  Details: {"shiftCount":8}
  
[OPERATION] 泉威國安班表解析完成
  Details: {"sheetName":"11501","totalRecords":21,"totalEmployees":3,"shiftCodes":"A, A1, B, B1, O"}
```

---

## 🎯 關鍵設計特點

### 1. **防禦性設計**
```javascript
try {
  // Log 寫入邏輯
} catch (error) {
  // 即使 Log 失敗，也不影響主程式
  Logger.log('記錄日誌失敗: ' + error.message);
}
```

### 2. **雙輸出機制**
- **Google Sheets** - 永久儲存，可查詢分析
- **Apps Script Logger** - 即時查看，開發除錯

### 3. **智慧過濾**
- OFF - 完全不記錄（效能最佳）
- OPERATION - 只記錄關鍵操作（節省空間）
- DEBUG - 記錄所有資訊（完整追蹤）

### 4. **自動化初始化**
- 第一次使用自動建立工作表
- 自動格式化標題列
- 自動設定欄寬

### 5. **結構化詳細資訊**
- 使用 JSON 格式儲存 details
- 便於程式化查詢和分析
- 支援複雜物件結構

---

## 💡 使用建議

### 開發階段
```javascript
// Script Properties 設定
Log_Level = 2  // DEBUG

// 可以看到所有細節
logDebug('變數值', { value: someValue });
logDebug('迴圈進度', { current: i, total: length });
```

### 測試階段
```javascript
// Script Properties 設定
Log_Level = 2  // DEBUG

// 驗證每個步驟
logOperation('測試開始', { testCase: 'case1' });
logDebug('中間結果', { result: intermediateResult });
logOperation('測試完成', { success: true });
```

### 正式環境（初期）
```javascript
// Script Properties 設定
Log_Level = 2  // DEBUG

// 觀察實際使用情況
logOperation('使用者上傳', { fileName });
logDebug('解析詳情', { details });
```

### 正式環境（穩定後）
```javascript
// Script Properties 設定
Log_Level = 1  // OPERATION

// 只記錄重要操作和錯誤
logOperation('檔案處理完成', { count: 100 });
logError('處理失敗', { error });
// logDebug 會被過濾掉
```

---

## 📊 Log 資料量預估

### 單筆記錄大小
```
時間: 20 bytes
等級: 10 bytes
訊息: 50 bytes (平均)
環境: 10 bytes
使用者: 30 bytes
詳細資訊: 200 bytes (平均)
─────────────────
總計: 320 bytes ≈ 0.32 KB
```

### 資料量預估
| 使用情境 | 每日 Log 數 | 每日大小 | 30 天大小 |
|---------|------------|---------|----------|
| DEBUG (開發) | ~1000 筆 | ~320 KB | ~9.6 MB |
| DEBUG (正式初期) | ~500 筆 | ~160 KB | ~4.8 MB |
| OPERATION (正式穩定) | ~100 筆 | ~32 KB | ~960 KB |
| OFF (效能測試) | 0 筆 | 0 KB | 0 KB |

**建議：** 每月執行 `cleanOldLogs(30)` 清理舊資料

---

## 🔍 Log 查詢技巧

### 在 Google Sheets 中查詢

#### 1. 查詢特定時間範圍的錯誤
```sql
=QUERY(Log!A:F, "SELECT * WHERE B='ERROR' AND A >= date '2026-02-01'", 1)
```

#### 2. 統計每日 Log 數量
```sql
=QUERY(Log!A:F, "SELECT A, COUNT(A) GROUP BY A LABEL COUNT(A) '數量'", 1)
```

#### 3. 查詢特定員工的處理記錄
```sql
=QUERY(Log!A:F, "SELECT * WHERE C CONTAINS 'TiNg'", 1)
```

#### 4. 只看 OPERATION 等級
```sql
=QUERY(Log!A:F, "SELECT * WHERE B='OPERATION' ORDER BY A DESC", 1)
```

---

**文件版本：** v1.0  
**建立時間：** 2026-02-06  
**對應專案版本：** v0.6
