# SalarySystem 環境設定指南

本文件說明如何設定薪資計算系統的開發與部署環境。

---

## 📋 前置需求

### 必要軟體
- Node.js（建議 v18 以上）
- npm 或 yarn
- Git
- Google 帳號（用於 GAS 與 Sheets）

### 選用軟體
- Clasp（Google Apps Script CLI）
- VSCode 或其他程式編輯器

---

## 🔧 步驟 1：安裝 Node.js 相依套件

```bash
cd SalarySystem
npm install
```

這會安裝 Jest 及其他測試相關套件。

---

## 🔧 步驟 2：安裝 Clasp

Clasp 是 Google Apps Script 的命令列工具，用於本地開發與同步。

### 全域安裝
```bash
npm install -g @google/clasp
```

### 登入 Google 帳號
```bash
clasp login
```

這會開啟瀏覽器，請登入您的 Google 帳號並授權 Clasp。

---

## 🔧 步驟 3：建立 Google Sheets

### 3.1 建立測試環境 Google Sheets

1. 前往 [Google Sheets](https://sheets.google.com)
2. 建立新試算表，命名為「薪資計算系統 - 測試」
3. 建立以下工作表：
   - 原始班表資料
   - 原始打卡資料
   - 計算結果（原始）
   - 計算結果（調整後）
   - 調整歷程記錄
   - 最終確認結果
   - 錯誤與異常記錄
   - 處理狀態
4. 記下試算表的 ID（網址中 `/d/` 後面的那串字元）

### 3.2 建立正式環境 Google Sheets

重複上述步驟，命名為「薪資計算系統 - 正式」。

---

## 🔧 步驟 4：建立 GAS 專案

### 4.1 建立測試環境 GAS 專案

#### 方法 A：透過 Clasp 建立（推薦）

```bash
cd gas-backend/test
clasp create --title "SalarySystem - Test" --type standalone
```

這會建立新的 GAS 專案，並自動產生 `.clasp.json`。

#### 方法 B：透過 Google Apps Script 網頁建立

1. 前往 [Google Apps Script](https://script.google.com)
2. 點擊「新專案」
3. 命名為「SalarySystem - Test」
4. 記下腳本 ID（檔案 → 專案屬性 → 腳本 ID）
5. 在本地建立 `.clasp.json`：

```json
{
  "scriptId": "你的腳本ID",
  "rootDir": "./gas-backend/test"
}
```

### 4.2 建立正式環境 GAS 專案

重複上述步驟，但：
- 命名為「SalarySystem - Production」
- `.clasp.json` 放在 `gas-backend/prod/`

---

## 🔧 步驟 5：設定 GAS Script Properties

### 5.1 測試環境

1. 開啟測試環境 GAS 專案
2. 點擊「專案設定」（左側齒輪圖示）
3. 在「指令碼屬性」區塊，新增以下屬性：

| 屬性名稱 | 值 |
|---------|---|
| `SHEET_ID` | 測試環境 Google Sheets ID |
| `ENVIRONMENT` | `test` |

### 5.2 正式環境

重複上述步驟，但：
- `SHEET_ID` 使用正式環境 Google Sheets ID
- `ENVIRONMENT` 設為 `production`

---

## 🔧 步驟 6：推送程式碼到 GAS

### 6.1 測試環境

```bash
cd gas-backend/test
npx clasp push
```

若出現「Manifest file has been updated. Do you want to push and overwrite?」，輸入 `y`。

### 6.2 正式環境

```bash
cd gas-backend/prod
npx clasp push
```

---

## 🔧 步驟 7：部署 GAS Web App

### 7.1 測試環境

1. 開啟測試環境 GAS 專案（`npx clasp open`）
2. 點擊「部署」→「新增部署」
3. 類型選擇「網頁應用程式」
4. 設定：
   - **說明**：測試環境初始部署
   - **執行身分**：我
   - **誰可以存取**：所有人（或依需求調整）
5. 點擊「部署」
6. 複製「網頁應用程式網址」（例如：`https://script.google.com/macros/s/.../exec`）
7. 更新 `frontend/config.js` 中的 `GAS_URL_TEST`

### 7.2 正式環境

重複上述步驟，更新 `GAS_URL_PROD`。

---

## 🔧 步驟 8：設定前端環境

編輯 `frontend/config.js`：

```javascript
const CONFIG = {
  GAS_URL_TEST: 'https://script.google.com/macros/s/測試環境ID/exec',
  GAS_URL_PROD: 'https://script.google.com/macros/s/正式環境ID/exec',
  ENVIRONMENT: 'test' // 開發時使用 'test'，部署時改為 'production'
};
```

---

## 🔧 步驟 9：執行測試

```bash
npm test
```

確保所有測試通過。

---

## 🔧 步驟 10：部署到 GitHub Pages

### 10.1 推送到 GitHub

```bash
git add .
git commit -m "Initial commit: project setup"
git push -u origin main
```

### 10.2 啟用 GitHub Pages

1. 前往 GitHub 專案頁面
2. 點擊「Settings」
3. 左側選單點擊「Pages」
4. Source 選擇「Deploy from a branch」
5. Branch 選擇 `main` 和 `/frontend`
6. 點擊「Save」

幾分鐘後，網站會部署在 `https://chyuanwei.github.io/SalarySystem/`。

---

## 🎉 完成！

現在您已經完成所有設定，可以開始開發了。

### 下一步

1. 確認薪資計算規則（見 `docs/CALCULATION.md`）
2. 確認 Excel 檔案格式（見 `docs/SCHEMA.md`）
3. 開始開發功能

---

## 🔍 常見問題

### Q: Clasp push 失敗，顯示「rootDir 不匹配」

**A**: 編輯 `.clasp.json`，確認 `rootDir` 指向正確的目錄路徑。

### Q: GAS 部署後無法存取

**A**: 檢查部署設定中的「誰可以存取」，確保設為「所有人」或適當的權限。

### Q: 前端無法呼叫 GAS API

**A**: 
1. 檢查 `frontend/config.js` 中的 GAS URL 是否正確
2. 檢查 GAS 部署設定中的 CORS 設定
3. 開啟瀏覽器開發者工具查看錯誤訊息

### Q: 測試失敗

**A**: 
1. 確認 `npm install` 已執行
2. 檢查 `jest.config.js` 設定
3. 查看錯誤訊息，可能是相依套件版本問題

---

**更多問題**請參考 [.cursor/CONTEXT.md](.cursor/CONTEXT.md) 或聯絡專案管理者。
