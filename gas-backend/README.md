# Google Apps Script 後端

本目錄包含薪資計算系統的 GAS 後端程式碼。

---

## 📂 目錄結構

```
gas-backend/
├── test/           # 測試環境 GAS
│   ├── .clasp.json # 測試環境腳本 ID（不提交 Git）
│   └── *.gs        # GAS 程式碼
│
└── prod/           # 正式環境 GAS
    ├── .clasp.json # 正式環境腳本 ID（不提交 Git）
    └── *.gs        # GAS 程式碼
```

---

## 🔑 GAS Script ID

### 測試環境
- **Script ID**: `1b_UcoqPQ8Vm6lekQli79JRePKDB-5qbbkLPh5lsfOdZOVtAD-Z61X8ZH`
- **專案名稱**: SalarySystem - Test
- **網址**: https://script.google.com/home/projects/1b_UcoqPQ8Vm6lekQli79JRePKDB-5qbbkLPh5lsfOdZOVtAD-Z61X8ZH

### 正式環境
- **Script ID**: `185v2ySyq9ncyzBwbnnogQTXHuw3S78pfmXcaATxWRs5K5W3ZYsSRG4F7`
- **專案名稱**: SalarySystem - Production
- **網址**: https://script.google.com/home/projects/185v2ySyq9ncyzBwbnnogQTXHuw3S78pfmXcaATxWRs5K5W3ZYsSRG4F7

---

## 🚀 開發流程

### 1. 測試環境開發

```bash
# 進入測試環境目錄
cd gas-backend/test

# 從 GAS 拉取最新程式碼
npx clasp pull

# 本地修改程式碼...

# 推送到 GAS
npx clasp push

# 開啟 GAS 編輯器
npx clasp open
```

### 2. 同步到正式環境

測試環境驗證通過後：

```bash
# 複製測試環境的 .gs 檔案到正式環境（排除 .clasp.json）
# Windows PowerShell:
Copy-Item gas-backend/test/*.gs gas-backend/prod/ -Exclude .clasp.json

# 進入正式環境目錄
cd gas-backend/prod

# 推送到正式 GAS
npx clasp push

# 開啟正式 GAS 編輯器
npx clasp open
```

---

## 📝 程式碼結構（規劃）

### 主要檔案

| 檔案名稱 | 說明 |
|---------|------|
| `Code.gs` | 主程式，包含 `doGet` 和 `doPost` |
| `Config.gs` | 環境設定與常數 |
| `ExcelParser.gs` | Excel 檔案解析 |
| `DataValidator.gs` | 資料驗證 |
| `SalaryCalculator.gs` | 薪資計算邏輯 |
| `SheetService.gs` | Google Sheets 操作 |
| `AdjustmentService.gs` | 調整功能 |
| `VersionControl.gs` | 版本管理 |
| `appsscript.json` | GAS 專案設定檔 |

---

## ⚙️ Script Properties 設定

### 測試環境需設定

在 GAS 專案 → 專案設定 → 指令碼屬性：

| 屬性名稱 | 值 | 說明 |
|---------|---|------|
| `SHEET_ID` | （測試用 Sheets ID） | 測試環境 Google Sheets ID |
| `ENVIRONMENT` | `test` | 環境標識 |

### 正式環境需設定

| 屬性名稱 | 值 | 說明 |
|---------|---|------|
| `SHEET_ID` | （正式用 Sheets ID） | 正式環境 Google Sheets ID |
| `ENVIRONMENT` | `production` | 環境標識 |

---

## 🌐 Web App 部署

### 測試環境部署

1. 開啟測試 GAS 專案：`cd gas-backend/test && npx clasp open`
2. 點擊「部署」→「新增部署」
3. 類型：網頁應用程式
4. 設定：
   - 說明：測試環境
   - 執行身分：我
   - 誰可以存取：所有人
5. 部署後複製網址，更新前端 `config.js` 的 `GAS_URL_TEST`

### 正式環境部署

重複上述步驟，更新 `GAS_URL_PROD`。

---

## 🔍 常見問題

### Q: clasp push 失敗，顯示 "rootDir 不匹配"

**A**: 編輯 `.clasp.json`，確認 `rootDir` 路徑正確：

```json
{
  "scriptId": "...",
  "rootDir": "c:\\IT-Project\\Chyuanwei\\SalarySystem\\gas-backend\\test"
}
```

### Q: 如何切換不同的 GAS 專案？

**A**: 確保在正確的目錄下執行 clasp 指令：
- 測試環境：`cd gas-backend/test`
- 正式環境：`cd gas-backend/prod`

### Q: .clasp.json 要提交到 Git 嗎？

**A**: 不要！`.clasp.json` 已列入 `.gitignore`，因為它包含本地路徑，每台電腦可能不同。

---

## 📌 注意事項

1. **測試優先**：所有新功能先在測試環境開發與驗證
2. **謹慎推送正式環境**：確保測試通過後才同步到正式環境
3. **記錄變更**：每次 GAS 更新後提供 Comment for deploy
4. **備份**：重要變更前先備份 GAS 專案

---

**更新日期**: 2024-02-06
