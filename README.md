# 薪資計算系統 (SalarySystem)

自動化薪資計算系統，比對班表與打卡紀錄，支援人工檢視與調整，最終確認並記錄至 Google Sheets。

---

## 📋 系統功能

- ✅ **檔案上傳**：透過網頁上傳班表與打卡紀錄 Excel 檔案
- ✅ **自動計算**：後端自動比對資料並計算薪資
- ✅ **檢視調整**：前端顯示計算結果，支援人工調整
- ✅ **歷程記錄**：所有調整動作記錄在 Google Sheets
- ✅ **二次確認**：確認前顯示調整摘要，避免誤操作
- ✅ **測試環境**：完整的測試與正式環境分離

---

## 🏗️ 系統架構

- **前端**：GitHub Pages（HTML/CSS/JavaScript）
- **後端**：Google Apps Script (GAS)
- **資料庫**：Google Sheets
- **測試**：Jest 單元測試與整合測試

---

## 📂 專案結構

```
SalarySystem/
├── .cursor/CONTEXT.md      # 專案脈絡文件（重要！）
├── frontend/               # 前端網頁
├── gas-backend/            # GAS 後端
│   ├── test/              # 測試環境
│   └── prod/              # 正式環境
├── tests/                  # Jest 測試
├── test-data/              # 測試資料
└── docs/                   # 文件
```

詳細結構請見 [.cursor/CONTEXT.md](.cursor/CONTEXT.md)。

---

## 🚀 快速開始

### 1. 複製專案

```bash
git clone https://github.com/chyuanwei/SalarySystem.git
cd SalarySystem
```

### 2. 安裝相依套件

```bash
npm install
```

### 3. 設定 GAS 環境

請參考 [SETUP.md](SETUP.md) 進行詳細設定。

### 4. 測試 GAS 連線

**重要**：在使用系統前，請先測試 GAS 連線是否正常。

開啟 `frontend/test/test-connection.html` 進行連線測試，確保：
- ✅ GAS 已部署
- ✅ GAS URL 設定正確
- ✅ Script Properties 已設定

### 5. 執行測試

```bash
npm test
```

---

## 📖 文件

- [專案 Context](.cursor/CONTEXT.md) - 完整的專案脈絡（**必讀**）
- [設定指南](SETUP.md) - 環境設定步驟
- [API 文件](docs/API.md) - API 端點說明
- [Excel 格式](docs/SCHEMA.md) - 班表與打卡紀錄格式定義
- [計算規則](docs/CALCULATION.md) - 薪資計算邏輯
- [調整功能](docs/ADJUSTMENT.md) - 調整功能說明
- [UI 規格](docs/UI_SPEC.md) - 前端介面規格

---

## 🔧 開發流程

1. **開發功能**：在 `gas-backend/test/` 進行開發
2. **執行測試**：`npm test` 確保測試通過
3. **推送測試環境**：`cd gas-backend/test && npx clasp push`
4. **驗證功能**：使用前端測試頁驗證
5. **同步正式環境**：複製到 `gas-backend/prod/` 並推送
6. **提交版控**：`git add . && git commit -m "..." && git push`

詳細流程請見 [CONTEXT.md](.cursor/CONTEXT.md) 第 9 節。

---

## 🧪 測試

### 單元測試
```bash
npm test
```

### 整合測試
```bash
npm run test:integration
```

### 測試覆蓋率
```bash
npm run test:coverage
```

---

## 🌐 部署

### 前端部署
GitHub Pages 會自動部署 `frontend/` 目錄的內容。入口為 `frontend/index.html`，可選擇測試環境（`/test/`）或正式環境（`/prod/`）。

### 後端部署
```bash
# 測試環境
cd gas-backend/test
npx clasp push

# 正式環境
cd gas-backend/prod
npx clasp push
```

---

## 📝 待辦事項

- [ ] 確認薪資計算規則
- [ ] 確認 Excel 檔案格式
- [ ] 完成前端 UI 設計
- [ ] 完成 GAS 後端開發
- [ ] 撰寫完整測試
- [ ] 建立測試資料集

---

## 🤝 貢獻

歡迎提交 Issue 或 Pull Request。

---

## 📄 授權

此專案為私人專案，未開放授權。

---

## 📞 聯絡

如有問題請聯絡專案管理者。

---

**通用開發慣例**請參考：`c:\IT-Project\.cursor\GENERAL_CONTEXT.md`
