# Log 系統說明

本文件說明薪資計算系統的 Log 架構與使用方式。

---

## 📋 Log 等級設定

### 等級定義

系統支援三個 Log 等級，透過 **指令碼屬性 `Log_Level`** 控制：

| 等級值 | 名稱 | 說明 | 記錄內容 |
|--------|------|------|----------|
| `0` | OFF | 關閉 Log | 完全不記錄任何 Log |
| `1` | OPERATION | 營運等級 | 只記錄重要的營運資訊與錯誤 |
| `2` | DEBUG | Debug 等級 | 記錄所有操作與詳細資訊（預設） |

---

## ⚙️ 設定方式

### 1. 在 GAS 專案中設定

1. 開啟 GAS 專案（測試或正式環境）
2. 點擊左側齒輪圖示「專案設定」
3. 在「指令碼屬性」區塊，新增或修改：

| 屬性名稱 | 值 | 說明 |
|---------|---|------|
| `Log_Level` | `0` / `1` / `2` | Log 等級設定 |

### 2. 建議設定

- **開發與測試階段**：設為 `2`（DEBUG），記錄所有資訊，方便除錯
- **正式環境初期**：設為 `2`（DEBUG），觀察系統運作狀況
- **正式環境穩定後**：設為 `1`（OPERATION），減少 Log 數量，只記錄重要資訊
- **不需要 Log 時**：設為 `0`（OFF），完全關閉

---

## 📝 Log 記錄內容

### Log Sheet 結構

所有 Log 都記錄在 Google Sheets 的 **`Log`** 工作表中，包含以下欄位：

| 欄位 | 說明 | 範例 |
|------|------|------|
| 時間 | Log 產生時間 | 2024-02-06 14:30:25 |
| 等級 | Log 等級 | DEBUG / INFO / WARNING / ERROR / OPERATION |
| 訊息 | 簡短的 Log 訊息 | 檔案處理成功: 2026泉威國安班表.xlsx |
| 環境 | 執行環境 | test / production |
| 使用者 | 執行者的 Email | user@example.com |
| 詳細資訊 | 額外的詳細資訊（JSON 格式） | `{"fileName":"...","rowCount":50}` |

### Log 等級說明

| 等級 | 使用時機 | 範例 |
|------|----------|------|
| **OPERATION** | 重要的營運資訊 | 檔案上傳成功、計算完成、系統啟動 |
| **ERROR** | 錯誤發生 | 檔案解析失敗、資料寫入錯誤、API 呼叫失敗 |
| **WARNING** | 警告訊息 | 資料格式不一致、欄位數量異常 |
| **INFO** | 一般資訊 | 資料驗證通過、工作表已建立 |
| **DEBUG** | 除錯資訊 | 詳細的執行步驟、變數值、函數呼叫 |

---

## 🔧 Log 等級與記錄規則

### Log_Level = 0 (OFF)
- ✘ 不記錄任何 Log
- 完全關閉 Log 系統

### Log_Level = 1 (OPERATION)
- ✓ 記錄 **OPERATION** 等級的 Log（重要營運資訊）
- ✓ 記錄 **ERROR** 等級的 Log（錯誤）
- ✘ 不記錄 DEBUG、INFO、WARNING

**適用場景**：正式環境穩定運作時，只關心重要操作與錯誤

### Log_Level = 2 (DEBUG)
- ✓ 記錄 **所有等級** 的 Log
- 包含：DEBUG、INFO、WARNING、ERROR、OPERATION

**適用場景**：開發、測試、除錯、正式環境初期監控

---

## 💻 程式中使用 Log

### 基本使用

```javascript
// 最通用的方式
logToSheet('這是一條訊息', 'INFO');

// 帶詳細資訊
logToSheet('檔案處理成功', 'INFO', {
  fileName: 'test.xlsx',
  rowCount: 100
});
```

### 便利函數

系統提供各等級的便利函數：

```javascript
// Debug 等級（Log_Level >= 2 時記錄）
logDebug('詳細的除錯資訊', { variable: value });

// Info 等級（Log_Level >= 2 時記錄）
logInfo('一般資訊', { data: something });

// Warning 等級（Log_Level >= 2 時記錄）
logWarning('警告訊息', { warning: details });

// Error 等級（Log_Level >= 1 時記錄）
logError('錯誤訊息', { error: errorObject });

// Operation 等級（Log_Level >= 1 時記錄）
logOperation('重要營運資訊', { operation: details });
```

### 使用範例

```javascript
function handleUpload(requestData) {
  // 記錄營運資訊（重要操作）
  logOperation(`開始處理上傳: ${fileName}`, {
    fileName: fileName,
    targetSheet: targetSheet
  });
  
  try {
    // 記錄除錯資訊
    logDebug(`檔案大小: ${fileSize} bytes`);
    
    // 處理邏輯...
    
    // 記錄一般資訊
    logInfo('資料驗證通過', {
      rowCount: rowCount
    });
    
    // 記錄警告
    if (hasWarnings) {
      logWarning('發現資料警告', {
        warnings: warningsList
      });
    }
    
    // 記錄營運成功
    logOperation('檔案處理完成', {
      processTime: processTime
    });
    
  } catch (error) {
    // 記錄錯誤
    logError('處理失敗', {
      error: error.toString(),
      stack: error.stack
    });
  }
}
```

---

## 🧹 Log 清理

### 自動清理舊 Log

系統提供 `cleanOldLogs()` 函數，可清理超過指定天數的舊 Log：

```javascript
// 清理 30 天前的 Log（預設）
cleanOldLogs();

// 清理 7 天前的 Log
cleanOldLogs(7);

// 清理 90 天前的 Log
cleanOldLogs(90);
```

### 設定自動清理

可以在 GAS 中設定觸發器，定期清理舊 Log：

1. 在 GAS 編輯器中，點擊左側「觸發條件」
2. 點擊「新增觸發條件」
3. 設定：
   - 選擇要執行的函式：`cleanOldLogs`
   - 選擇活動來源：時間驅動
   - 選擇時間型觸發條件類型：每週計時器
   - 選擇星期幾：星期日
   - 選擇時段：凌晨 2 時至 3 時
4. 儲存

---

## 📊 Log 分析

### 查看 Log

1. 開啟 Google Sheets
2. 切換到 **`Log`** 工作表
3. 使用 Google Sheets 的篩選功能：
   - 篩選特定等級（例如：只看 ERROR）
   - 篩選特定時間範圍
   - 搜尋特定關鍵字

### 常見查詢

#### 查看所有錯誤
1. 點擊「等級」欄位的篩選按鈕
2. 只勾選「ERROR」
3. 點擊「確定」

#### 查看特定檔案的處理記錄
1. 使用 Ctrl+F 搜尋檔案名稱
2. 查看相關的 Log 記錄

#### 查看今天的 Log
1. 點擊「時間」欄位的篩選按鈕
2. 選擇「條件」→「日期是」→「今天」

---

## 🎯 最佳實踐

### 1. 適當使用 Log 等級

```javascript
// ✅ 好的做法
logOperation('使用者上傳檔案', { fileName: 'test.xlsx' });  // 重要操作
logError('檔案解析失敗', { error: error.message });        // 錯誤
logWarning('資料格式異常', { issue: 'missing column' });    // 警告
logInfo('資料驗證完成', { rowCount: 100 });                // 一般資訊
logDebug('進入函數 parseExcel', { params: {...} });        // 除錯

// ❌ 不好的做法
logOperation('變數值為 123');  // 這應該是 DEBUG
logDebug('檔案上傳成功');      // 這應該是 OPERATION
```

### 2. 提供有用的詳細資訊

```javascript
// ✅ 好的做法
logError('檔案處理失敗', {
  fileName: fileName,
  error: error.toString(),
  stack: error.stack,
  step: 'parsing'
});

// ❌ 不好的做法
logError('失敗');  // 沒有足夠資訊
```

### 3. 記錄處理時間

```javascript
const startTime = new Date();
// ... 處理邏輯 ...
const endTime = new Date();
const processTime = (endTime - startTime) / 1000;

logOperation('處理完成', {
  processTime: `${processTime.toFixed(2)} 秒`
});
```

### 4. 不要過度 Log

```javascript
// ❌ 不好的做法（太多 Debug Log）
for (let i = 0; i < 1000; i++) {
  logDebug(`處理第 ${i} 筆`);  // 會產生 1000 筆 Log！
}

// ✅ 好的做法
logDebug(`開始處理 ${totalCount} 筆資料`);
// ... 處理邏輯 ...
logDebug(`處理完成，成功 ${successCount} 筆，失敗 ${failCount} 筆`);
```

---

## 🔍 疑難排解

### 問題 1：Log 沒有寫入

**可能原因**：
- `Log_Level` 設為 `0`（OFF）
- Log 等級設定不正確（例如：OPERATION 等級時使用 logDebug）

**解決方式**：
1. 檢查 Script Properties 中的 `Log_Level` 設定
2. 確認使用的 Log 函數等級是否會被記錄

### 問題 2：Log 過多導致效能問題

**解決方式**：
1. 將 `Log_Level` 從 `2` 改為 `1`
2. 定期執行 `cleanOldLogs()` 清理舊 Log
3. 檢視程式碼，減少不必要的 Debug Log

### 問題 3：Log Sheet 格式亂掉

**解決方式**：
1. 刪除整個 `Log` 工作表
2. 重新執行任何會記錄 Log 的操作
3. 系統會自動重建並格式化 Log Sheet

---

## 📌 注意事項

1. **效能影響**：
   - Log 等級 `2` (DEBUG) 會記錄大量資訊，可能略微影響效能
   - 正式環境穩定後建議改為等級 `1` (OPERATION)

2. **儲存空間**：
   - Log 會累積在 Google Sheets 中
   - 建議定期清理舊 Log（例如：保留最近 30 天）

3. **隱私考量**：
   - Log 中可能包含敏感資訊（檔案名稱、使用者 Email）
   - 確保 Google Sheets 的存取權限設定正確

4. **預設值**：
   - 如果未設定 `Log_Level`，預設為 `2` (DEBUG)
   - 建議明確設定以避免混淆

---

**更新日期**: 2024-02-06
