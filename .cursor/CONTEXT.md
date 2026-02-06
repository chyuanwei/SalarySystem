# SalarySystem 專案 Context

本檔案為此專案專用脈絡，供 AI 與開發者快速掌握結構、流程與慣例。  
**通用慣例**（Git、clasp、Jest、分析後先問再改等）見：`c:\IT-Project\.cursor\GENERAL_CONTEXT.md`。

---

## 1. 專案概述

### 系統目的
計算員工薪資，比對班表與打卡紀錄，支援人工檢視與調整，最終確認並記錄至 Google Sheets。

### 系統架構
- **前端**：GitHub Pages 網頁，供使用者上傳 Excel 檔案、檢視計算結果、調整薪資、確認送出。
- **後端**：Google Apps Script (GAS)，負責檔案解析、資料驗證、薪資計算、資料儲存。
- **資料儲存**：Google Sheets，記錄原始資料、計算結果、調整歷程、最終確認結果。

### 使用者流程
1. **上傳**：使用者透過網頁上傳兩個 Excel 檔案（班表、打卡紀錄）
2. **處理**：GAS 後端解析檔案、驗證資料、計算薪資
3. **檢視與調整**：前端顯示計算結果，使用者可檢視明細並進行調整（需填寫調整理由）
4. **確認**：使用者確認結果（二次確認機制），資料寫入最終結果表
5. **修改**：確認後仍可返回修改

---

## 2. 目錄結構

```
SalarySystem/
├── .cursor/
│   ├── CONTEXT.md              ← 本檔案
│   └── rules/
│       └── gas-deploy-comment.mdc
├── .gitignore
├── README.md
├── SETUP.md
│
├── frontend/                   # GitHub Pages 網頁
│   ├── index.html              # 主頁：上傳檔案
│   ├── processing.html         # 處理狀態頁面
│   ├── review.html             # 檢視與調整頁面
│   ├── complete.html           # 完成頁面
│   ├── styles/
│   │   ├── main.css
│   │   └── review.css
│   ├── scripts/
│   │   ├── upload.js           # 上傳邏輯
│   │   ├── review.js           # 檢視與調整邏輯
│   │   ├── api.js              # API 呼叫封裝
│   │   └── utils.js
│   └── config.js               # 環境設定（測試/正式 GAS URL）
│
├── gas-backend/                # Google Apps Script 後端
│   ├── test/                   # 測試環境 GAS
│   │   ├── .clasp.json         # 測試環境腳本 ID（不提交）
│   │   ├── .claspignore
│   │   ├── appsscript.json
│   │   ├── Code.gs             # 主程式（doGet/doPost）
│   │   ├── Config.gs           # 環境設定
│   │   ├── ExcelParser.gs     # Excel 解析
│   │   ├── DataValidator.gs   # 資料驗證
│   │   ├── SalaryCalculator.gs # 薪資計算邏輯
│   │   ├── SheetService.gs    # Google Sheets 操作
│   │   ├── AdjustmentService.gs # 調整邏輯
│   │   └── VersionControl.gs  # 版本管理
│   │
│   └── prod/                   # 正式環境 GAS
│       ├── .clasp.json         # 正式環境腳本 ID（不提交）
│       ├── .claspignore
│       └── (與 test 相同的 .gs 檔案)
│
├── tests/                      # 測試
│   ├── jest.config.js
│   ├── jest.setup.js
│   ├── unit/
│   │   ├── excelParser.test.js
│   │   ├── dataValidator.test.js
│   │   ├── salaryCalculator.test.js
│   │   └── adjustment.test.js
│   └── integration/
│       └── fullFlow.test.js
│
├── test-data/                  # 測試資料
│   ├── sample-schedule.xlsx    # 範例班表
│   ├── sample-attendance.xlsx  # 範例打卡紀錄
│   └── expected-result.json    # 預期計算結果
│
├── docs/                       # 文件
│   ├── API.md                  # API 端點說明
│   ├── SCHEMA.md               # Excel 格式定義
│   ├── CALCULATION.md          # 薪資計算規則
│   ├── ADJUSTMENT.md           # 調整功能文件
│   └── UI_SPEC.md              # UI 規格文件
│
└── package.json                # npm 相關（Jest 等）
```

---

## 3. 資料流（重要）

### 3.1 上傳與處理流程

```
使用者 → 上傳 Excel (班表 + 打卡紀錄)
       ↓
GAS doPost → /upload
       ↓
1. 檔案驗證（格式、大小、必要欄位）
2. 解析 Excel → JSON
3. 資料驗證（日期格式、時間格式、異常值）
4. 寫入「原始班表資料」、「原始打卡資料」表
5. 比對班表與打卡
6. 計算薪資
7. 寫入「計算結果（原始）」表
8. 回傳 jobId 與狀態
       ↓
前端輪詢 → GAS doGet → /status/:jobId
       ↓
處理完成 → 前端導向 review.html
```

### 3.2 檢視與調整流程

```
前端 → GAS doGet → /result/:jobId
     ↓
取得計算結果（原始）
     ↓
前端顯示 → 使用者檢視明細
     ↓
使用者調整（填寫理由）→ 前端暫存
     ↓
使用者點擊「確認」→ 前端 POST → GAS doPut → /adjust/:jobId
     ↓
GAS 寫入「計算結果（調整後）」+ 「調整歷程記錄」
     ↓
回傳成功 → 前端顯示二次確認對話框
     ↓
使用者確認 → 前端 POST → GAS doPost → /confirm/:jobId
     ↓
GAS 寫入「最終確認結果」
     ↓
前端導向 complete.html
```

### 3.3 Excel 檔案格式

#### 班表 Excel（必要欄位）
- 員工編號
- 員工姓名
- 日期
- 班別（例如：早班、晚班）
- 預定工作時數

#### 打卡紀錄 Excel（必要欄位）
- 員工編號
- 日期
- 上班打卡時間
- 下班打卡時間

---

## 4. API 端點設計

所有端點透過 GAS Web App 部署，使用 `doGet` 和 `doPost` 處理。

| 方法 | 路徑 | 說明 | 參數 |
|------|------|------|------|
| POST | `/upload` | 上傳 Excel 檔案 | `scheduleFile`, `attendanceFile` (base64) |
| GET | `/status/:jobId` | 查詢處理狀態 | `jobId` |
| GET | `/result/:jobId` | 取得計算結果 | `jobId` |
| PUT | `/adjust/:jobId` | 提交調整資料 | `jobId`, `adjustments` (陣列) |
| POST | `/confirm/:jobId` | 確認最終結果 | `jobId` |
| GET | `/export/:jobId` | 匯出報表 | `jobId`, `format` (xlsx/pdf) |

---

## 5. Google Sheets 表結構

專案使用一個 Google Sheets 檔案，包含以下工作表：

| 工作表名稱 | 說明 | 主要欄位 |
|------------|------|----------|
| **原始班表資料** | 上傳的班表 Excel 內容 | 員工編號、姓名、日期、班別、預定時數 |
| **原始打卡資料** | 上傳的打卡紀錄 Excel 內容 | 員工編號、日期、上班時間、下班時間 |
| **計算結果（原始）** | 系統計算的原始薪資結果 | 員工編號、姓名、工作天數、總時數、正常時數、加班時數、薪資明細、總薪資 |
| **計算結果（調整後）** | 使用者調整後的薪資結果 | 同上 + 調整欄位標記 |
| **調整歷程記錄** | 記錄所有調整動作 | 時間戳記、jobId、員工編號、調整欄位、原始值、調整後值、調整理由、操作者 |
| **最終確認結果** | 確認後的最終版本 | 同「計算結果（調整後）」+ 確認時間、確認者 |
| **錯誤與異常記錄** | 記錄處理過程中的錯誤 | 時間戳記、jobId、錯誤類型、錯誤訊息、相關資料 |
| **處理狀態** | 記錄每個 job 的處理狀態 | jobId、上傳時間、狀態、完成時間、錯誤訊息 |

---

## 6. 薪資計算邏輯（待定義）

> **注意**：以下為預設邏輯，實際計算規則需與使用者確認。

### 6.1 基本假設
- 時薪制
- 正常工時：每日 8 小時
- 加班費率：
  - 平日加班前 2 小時：1.34x
  - 平日加班超過 2 小時：1.67x
  - 例假日：2x

### 6.2 計算步驟
1. 比對班表與打卡：找出實際工作時間
2. 計算正常工時與加班工時
3. 套用費率計算薪資
4. 處理異常：遲到、早退、缺卡

### 6.3 異常處理
- **缺卡**：有班表但無打卡 → 標記異常，人工判斷
- **遲到/早退**：實際工時少於班表 → 計算實際工時
- **加班**：實際工時多於班表 → 計算加班時數

---

## 7. 測試策略

### 7.1 單元測試
- Excel 解析函數
- 資料驗證函數
- 薪資計算邏輯函數
- 調整功能函數

### 7.2 整合測試
- 完整流程測試（上傳 → 計算 → 調整 → 確認）
- 異常情況測試（缺卡、格式錯誤等）
- 邊界值測試

### 7.3 測試工具
- Jest（JavaScript 測試框架）
- Clasp（本地開發 GAS）
- 測試資料集

---

## 8. 環境分離

### 8.1 測試環境
- **GAS 專案**：SalarySystem - Test
- **Script ID**：`1b_UcoqPQ8Vm6lekQli79JRePKDB-5qbbkLPh5lsfOdZOVtAD-Z61X8ZH`
- **Script Properties**：
  - `SHEET_ID`: 測試用 Google Sheets ID
  - `ENVIRONMENT`: `test`
  - `Log_Level`: `2`（建議，Debug 等級）
- **前端設定**：`config.js` 中 `GAS_URL_TEST`

### 8.2 正式環境
- **GAS 專案**：SalarySystem - Production
- **Script ID**：`185v2ySyq9ncyzBwbnnogQTXHuw3S78pfmXcaATxWRs5K5W3ZYsSRG4F7`
- **Script Properties**：
  - `SHEET_ID`: 正式用 Google Sheets ID
  - `ENVIRONMENT`: `production`
  - `Log_Level`: `1`（建議，營運等級，只記錄重要操作與錯誤）
- **前端設定**：`config.js` 中 `GAS_URL_PROD`

### 8.3 開發流程

#### 日常開發（使用 `go` 指令）
1. 在 `gas-backend/test/` 開發並測試
2. 執行 `go` 指令：
   - 自動 commit & push GitHub
   - 自動 `clasp push` 到測試環境
   - 自動更新 CONTEXT.md
3. 使用前端測試頁驗證功能

#### 部署到正式環境（需明確指令）
1. 測試環境驗證完成後，**明確告知 AI** 要部署到正式環境
2. AI 執行 `deploy-prod`：
   - 複製 `gas-backend/test/*.gs` 到 `gas-backend/prod/`（排除 `.clasp.json`）
   - `cd gas-backend/prod && clasp push`
   - Commit 並 push 到 GitHub
   - 更新部署記錄

---

## 9. 開發與部署

### 9.1 GitHub
- **遠端**：`origin` → https://github.com/chyuanwei/SalarySystem.git
- **分支策略**：
  - `main`：正式版本
  - `develop`：開發版本
  - `feature/*`：功能分支

### 9.2 GAS（clasp）
- **測試環境**：
  - 腳本 ID 存在 `gas-backend/test/.clasp.json`（不提交）
  - `cd gas-backend/test && npx clasp pull/push`
- **正式環境**：
  - 腳本 ID 存在 `gas-backend/prod/.clasp.json`（不提交）
  - `cd gas-backend/prod && npx clasp pull/push`

### 9.3 Jest
- 執行目錄：專案根目錄（內有 `tests/jest.config.js`）
- 指令：`npm test` 或 `npx jest`
- 測試前需先 `npm install`

---

## 10. 專案慣例

| 項目 | 說明 |
|------|------|
| **go** | 等同 **執行 + 更新（僅測試環境）**：執行當下需求（改 code、測試等）→ commit/push GitHub **（需附 commit message）** → **clasp push 測試環境 GAS** → 更新 `.cursor/CONTEXT.md` → 提供 **Comment for GAS deploy**（精簡版）。**重要**：(1) `go` 指令只推送測試環境，正式環境需另外指令。(2) **僅使用者可下 `go` 指令，AI 不可主動執行**。 |
| **deploy-prod** | 部署到正式環境：複製 `gas-backend/test/*.gs` 到 `gas-backend/prod/` → `cd gas-backend/prod && clasp push` → commit/push GitHub **（需附 commit message）** → 更新部署記錄 → 提供 **Comment for GAS deploy**。**僅在使用者明確指示時執行**。 |
| **Git commit** | 所有 `git commit` 必須附上有意義的 commit message，說明變更內容。不可使用過於簡略或無意義的訊息。 |
| **GAS 更新** | 凡修改 `gas-backend/**/*.gs`，回覆結尾須附 **Comment for GAS deploy**（精簡版，1-3 行，供 GAS 部署版本說明使用）。 |
| **分析後先問再改** | 分析完問題或提出解法後，**須先詢問使用者意見**，經同意後才執行修改；不得直接進行程式修改。 |
| **.clasp.json** | 已列入 .gitignore，勿提交；測試與正式環境各有各自的腳本 ID。 |
| **測試優先** | 新增功能時先寫測試，確保程式正確性。 |
| **文件同步更新** | 修改功能時同步更新相關文件（API.md、CALCULATION.md 等）。 |

---

## 11. 調整功能細節

### 11.1 調整權限
- 目前不限制權限，所有使用者皆可調整
- 未來可擴充：透過登入機制區分角色（操作者、主管）

### 11.2 調整記錄必填欄位
- 調整理由（必填）
- 調整時間（自動記錄）
- 調整者（目前無登入機制，可留空或未來擴充）

### 11.3 調整歷程保存
- 所有調整記錄永久保存在「調整歷程記錄」表
- 可追溯任何一筆調整的完整歷史

### 11.4 二次確認機制
- 使用者點擊「確認」後，顯示對話框列出調整摘要
- 使用者再次確認後才寫入最終結果
- 確認後仍可返回修改（重新進入調整流程）

---

## 12. 已實作功能（v0.1）

### 12.1 檔案上傳與解析
- ✅ 前端支援拖曳或選擇 Excel 檔案上傳
- ✅ 自動轉換為 Base64 編碼
- ✅ 檔案大小與格式驗證
- ✅ 上傳進度顯示
- ✅ 可自訂要處理的 Excel 工作表名稱
- ✅ 完整的 RWD 響應式設計（手機、平板、桌面優化）

### 12.2 Excel 解析
- ✅ 支援 .xlsx 和 .xls 格式
- ✅ 可指定要解析的工作表名稱（例如：11501）
- ✅ 自動轉換 Excel 為 Google Sheets
- ✅ 資料格式驗證

### 12.3 資料儲存
- ✅ 寫入指定的 Google Sheets 工作表
- ✅ 自動格式化標題列（藍底白字、粗體）
- ✅ 自動調整欄寬
- ✅ 凍結標題列

### 12.4 Log 系統
- ✅ 完整的 Log 架構（DEBUG / INFO / WARNING / ERROR / OPERATION）
- ✅ 透過 Script Properties `Log_Level` 控制 Log 等級（0=關閉, 1=營運, 2=Debug）
- ✅ 記錄到 Google Sheets 的 `Log` 工作表
- ✅ 包含時間、等級、訊息、環境、使用者、詳細資訊
- ✅ 便利函數：logDebug(), logInfo(), logWarning(), logError(), logOperation()
- ✅ 自動清理舊 Log 功能（cleanOldLogs）
- ✅ 詳細的 Log 說明文件（docs/LOG_SYSTEM.md）

### 12.5 測試與正式環境
- ✅ 完整的環境分離
- ✅ 獨立的 GAS Script ID
- ✅ 獨立的 Google Sheets

---

## 13. 待實作功能

### 13.1 薪資計算
- ☐ 比對班表與打卡紀錄
- ☐ 計算正常工時與加班工時
- ☐ 套用薪資計算規則
- ☐ 處理異常情況（缺卡、遲到、早退）

### 13.2 調整功能
- ☐ 前端檢視計算結果
- ☐ 支援人工調整
- ☐ 調整理由記錄
- ☐ 二次確認機制

### 13.3 報表匯出
- ☐ 匯出 Excel 格式
- ☐ 匯出 PDF 格式

---

## 14. 待確認項目

以下項目需與使用者進一步確認：

1. **薪資計算規則**：
   - 時薪/月薪制？
   - 加班費率？
   - 特殊津貼或扣款？

2. **Excel 格式**：
   - 班表的確切欄位名稱與順序
   - 打卡紀錄的確切欄位名稱與順序
   - 日期與時間格式

3. **異常處理規則**：
   - 缺卡如何處理？
   - 遲到早退如何扣款？

4. **報表需求**：
   - 需要哪些報表？
   - 匯出格式（Excel/PDF）？

---

## 13. 附錄：曾遇問題與注意

| 狀況 | 說明／解法 |
|------|------------|
| **Git index.lock** | 若 `git status` 報 index.lock 錯誤，多為上次操作中斷；可關閉其他 Git 程式後刪除 `.git/index.lock`。 |
| **clasp push 失敗（rootDir）** | 錯誤提示 rootDir 不匹配時，編輯 `.clasp.json`，將 `rootDir` 改為實際目錄路徑。 |
| **測試與正式環境混淆** | 務必確認當前操作的是 `gas-backend/test/` 還是 `gas-backend/prod/`，避免誤推。 |
| **Excel 解析錯誤** | 確認 Excel 檔案格式為 .xlsx 或 .xls，且欄位名稱與預期一致。 |
| **GAS 執行逾時** | GAS 單次執行限制 6 分鐘，若資料量大需考慮分批處理。 |

---

## 15. 部署記錄

| 日期 | 版本 | 環境 | 更新內容 | 部署者 |
|------|------|------|----------|--------|
| 2024-02-06 | v0.2 | 測試+正式 | 新增完整 Log 系統，支援 Log_Level 控制（0/1/2），五種 Log 等級，記錄到 'Log' sheet | AI |
| 2024-02-06 | v0.1 | 測試+正式 | 初始版本：Excel 上傳、解析指定工作表、寫入 Google Sheets、基本 Log 功能 | AI |

---

*本檔案為專案專用 context，請隨重要變更更新。*
*最後更新：2024-02-06*
