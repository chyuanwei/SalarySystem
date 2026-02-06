# 部署與測試指南

本文件說明如何部署與測試薪資計算系統。

---

## 📋 前置準備

### 1. 建立 Google Sheets

1. 前往 [Google Sheets](https://sheets.google.com)
2. 建立新試算表，命名為 `SalarySystemTest`
3. 記下試算表 ID（網址列 `/d/` 後面的字串）

### 2. 設定 Script Properties

1. 開啟測試 GAS 專案：
   ```bash
   cd gas-backend/test
   npx clasp open
   ```

2. 點擊左側齒輪圖示「專案設定」

3. 在「指令碼屬性」區塊，新增以下屬性：

| 屬性名稱 | 值 | 說明 |
|---------|---|------|
| `SHEET_ID` | `YOUR_SHEET_ID` | 步驟 1 記下的試算表 ID |

---

## 🚀 部署步驟

### 測試環境部署

#### 步驟 1：推送程式碼到 GAS

```bash
cd gas-backend/test
npx clasp push
```

#### 步驟 2：啟用 Drive API

1. 在 GAS 編輯器中，點擊左側「服務」
2. 搜尋「Google Drive API」
3. 選擇「Drive API v3」
4. 點擊「新增」

#### 步驟 3：部署為 Web App

1. 點擊右上角「部署」→「新增部署」
2. 類型選擇「網頁應用程式」
3. 設定：
   - **說明**：測試環境初始部署
   - **執行身分**：我
   - **誰可以存取**：所有人
4. 點擊「部署」
5. 複製「網頁應用程式網址」

#### 步驟 4：更新前端設定

編輯 `frontend/config.js`，將步驟 3 複製的網址更新到 `GAS_URL_TEST`：

```javascript
const CONFIG = {
  GAS_URL_TEST: 'https://script.google.com/macros/s/YOUR_DEPLOYMENT_ID/exec',
  // ...
};
```

---

## 🧪 測試流程

### 測試 1：測試 GAS 連線

1. 在 GAS 編輯器中，選擇函數 `testUpload`
2. 點擊「執行」
3. 檢查執行記錄：
   - 應顯示「Google Sheets 連線成功」
   - 應顯示「測試完成」

### 測試 2：測試前端連線

1. 開啟 `frontend/index.html`（可用瀏覽器直接開啟）
2. 開啟瀏覽器開發者工具（F12）
3. 在 Console 輸入：
   ```javascript
   fetch(CONFIG.GAS_URL + '?action=test')
     .then(r => r.json())
     .then(d => console.log(d));
   ```
4. 應該看到回應：`{success: true, message: '測試連線成功', environment: 'test'}`

### 測試 3：測試檔案上傳

1. 開啟 `frontend/index.html`
2. 選擇測試檔案（例如：`test-data/schedule-sample.xlsx`）
3. 點擊「開始上傳並處理」
4. 觀察進度條與訊息
5. 完成後，檢查 Google Sheets：
   - 應該有新的工作表「國安班表」
   - 工作表中應該有從 Excel 讀取的資料
   - 應該有新的工作表「處理記錄」，記錄處理歷程

### 測試 4：檢查日誌

1. 開啟 Google Sheets `SalarySystemTest`
2. 查看「處理記錄」工作表
3. 應該看到所有處理記錄，包含：
   - 時間戳記
   - 檔案名稱
   - 來源工作表
   - 目標工作表
   - 資料筆數
   - 狀態（SUCCESS/ERROR）

---

## 🐛 常見問題排除

### 問題 1：GAS 連線失敗

**症狀**：前端無法連線到 GAS

**解決方式**：
1. 確認 `frontend/config.js` 中的 `GAS_URL_TEST` 正確
2. 確認 GAS 已部署為 Web App
3. 確認「誰可以存取」設為「所有人」

### 問題 2：Drive API 錯誤

**症狀**：上傳時出現「Drive API 未啟用」錯誤

**解決方式**：
1. 在 GAS 編輯器中啟用 Drive API（參考部署步驟 2）
2. 重新部署 Web App

### 問題 3：找不到工作表

**症狀**：錯誤訊息「找不到工作表 "11501"」

**解決方式**：
1. 確認上傳的 Excel 檔案中確實有名為「11501」的工作表
2. 或修改 `frontend/config.js` 中的 `TARGET_SHEET_NAME` 為實際的工作表名稱

### 問題 4：SHEET_ID 錯誤

**症狀**：錯誤訊息「Google Sheets ID 設定錯誤」

**解決方式**：
1. 確認 Script Properties 中的 `SHEET_ID` 設定正確
2. 確認該 Google Sheets 存在且有存取權限

### 問題 5：CORS 錯誤

**症狀**：瀏覽器 Console 出現 CORS 相關錯誤

**說明**：這是正常的！前端使用 `no-cors` 模式發送請求，無法讀取回應內容，但請求會成功執行。

**驗證方式**：直接檢查 Google Sheets 是否有新資料。

---

## 📝 正式環境部署

測試環境驗證通過後，可部署到正式環境：

### 步驟 1：設定正式環境 Script Properties

1. 開啟正式 GAS 專案：
   ```bash
   cd gas-backend/prod
   npx clasp open
   ```

2. 設定 Script Properties：
   - `SHEET_ID`：正式環境的 Google Sheets ID
   - `ENVIRONMENT`：`production`

### 步驟 2：推送並部署

```bash
cd gas-backend/prod
npx clasp push
```

然後按照測試環境的部署步驟 2-3 進行。

### 步驟 3：更新前端設定

編輯 `frontend/config.js`：

```javascript
const CONFIG = {
  // ...
  ENVIRONMENT: 'production', // 改為 production
  // ...
};
```

---

## ✅ 驗收標準

系統正常運作應滿足：

- ✅ 前端可成功上傳 Excel 檔案
- ✅ GAS 可解析 Excel 中指定的工作表
- ✅ 資料正確寫入 Google Sheets
- ✅ 處理記錄完整記錄所有操作
- ✅ 錯誤訊息清楚易懂
- ✅ 標題列格式化正確（藍底白字）
- ✅ 欄寬自動調整

---

## 📞 支援

如有問題，請參考：
- [專案 Context](.cursor/CONTEXT.md)
- [API 文件](docs/API.md)
- [GAS 後端說明](gas-backend/README.md)

---

**更新日期**: 2024-02-06
