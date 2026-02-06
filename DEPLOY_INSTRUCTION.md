# v0.6 部署指引

## ✅ 已完成

### 1. Git 提交與推送
- ✅ Commit: `5ff37ad` - v0.6: 替換為泉威國安班表專用解析器
- ✅ Pushed to GitHub: https://github.com/chyuanwei/SalarySystem

### 2. 變更摘要
- 18 個檔案變更
- +3984 行新增
- -813 行刪除

## 🚀 待執行：部署到 GAS

由於 `.clasp.json` 已在 `.gitignore` 中排除（避免洩漏專案 ID），需要手動執行以下步驟：

### 測試環境部署

```bash
cd gas-backend/test
npx clasp push
```

### 📝 GAS 部署註解（Deploy Comment）

請在 GAS Editor 執行 Deploy 時使用以下註解：

```
v0.6: 替換為泉威國安班表專用解析器

✨ 新功能：
- 代碼式班別自動掃描（A,B,O,C等）
- 橫向日期佈局自動定位
- 手寫時間格式解析
- 自動計算工作時數（含跨日處理）
- 年月份自動偵測

🔧 技術變更：
- 移除通用欄位映射邏輯（~270行）
- 新增泉威國安班表專用解析器（~240行）
- 新增 TestUtils.gs 測試工具

📦 核心函數：
- parseQuanWeiSchedule() - 主解析
- buildShiftMap() - 代碼字典掃描
- locateDataStructure() - 結構定位
- parseTimeRange() - 時間解析與時數計算

✅ 本地驗證：通過
🧪 待測試：GAS環境測試、實際檔案上傳
```

## 🧪 測試步驟

### 第一階段：GAS Editor 測試

1. **環境檢查**
   - 打開 GAS Editor: `gas-backend/test/` 專案
   - 執行：`checkEnvironmentSetup()`
   - 確認：`SHEET_ID` 和 `Log_Level` 已設定

2. **單元測試**
   - 執行：`testParseQuanWeiSchedule()`
   - 檢查 Apps Script Logger 輸出
   - 預期：✅ 所有測試通過

### 第二階段：實際檔案上傳測試

1. **準備**
   - 檔案：`2026泉威國安班表.xlsx`
   - 工作表：`11501`, `11502`

2. **測試流程**
   - 打開 `frontend/index.html`
   - 上傳檔案，選擇 `11501`
   - 上傳檔案，選擇 `11502`

3. **驗證結果**
   - 檢查「國安班表」工作表：資料是否正確寫入
   - 檢查「Log」工作表：是否有完整記錄
   - 驗證欄位：員工姓名、排班日期、上班時間、下班時間、工作時數
   - 驗證格式：日期 YYYY/MM/DD、時間 HH:MM、時數 小數點1位

### 第三階段：生產環境部署（測試通過後）

```bash
cd gas-backend/prod
npx clasp push
```

## 📊 部署檢查清單

- [x] Git commit 完成
- [x] Git push 完成
- [ ] GAS test 環境 push
- [ ] GAS Editor 單元測試
- [ ] 實際檔案上傳測試
- [ ] 生產環境部署
- [ ] 更新 CONTEXT.md 部署記錄

## 📚 相關文件

- `HOW_TO_TEST.md` - 詳細測試指南
- `COMPLETION_CHECKLIST.md` - 完整檢查清單
- `REPLACEMENT_SUMMARY.md` - 替換摘要
- `LOG_SYSTEM_ANALYSIS.md` - Log系統分析

## 🔗 GitHub

Commit: https://github.com/chyuanwei/SalarySystem/commit/5ff37ad

---

**下一步：執行 GAS 環境測試**
