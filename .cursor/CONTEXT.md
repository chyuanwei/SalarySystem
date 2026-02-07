# SalarySystem 專案 Context

本檔案為此專案專用脈絡，供 AI 與開發者快速掌握結構、流程與慣例。  
**通用慣例**（Git、clasp、Jest、分析後先問再改等）見：`c:\IT-Project\.cursor\GENERAL_CONTEXT.md` 或 `it-project-cursor/.cursor/GENERAL_CONTEXT.md`。

---

## 1. 專案概述

### 系統目的
處理泉威國安班表，自動解析代碼式排班資料（A, B, O 等班別代碼），轉換為標準化排班記錄，並儲存至 Google Sheets。另支援查詢現行國安班表資料。

### 系統架構
- **前端**：入口頁 `index.html` 選擇 test/prod 環境；`frontend/test/`、`frontend/prod/` 各自含**查詢國安班表**與**上傳**功能。
- **後端**：Google Apps Script (GAS)，負責 Excel 解析、班表/打卡寫入、loadSchedule 查詢 API。
- **資料儲存**：Google Sheets（國安班表、打卡紀錄、Log、處理記錄）。

### 使用者流程

**查詢國安班表**
1. 選擇日期篩選（整月 YYYYMM 或單日）
2. 載入 → 取得該範圍資料與人員名單
3. 可多選人員再載入（AND 篩選）
4. 顯示結果卡片

**上傳**
1. 選擇上傳類型：班表 | 打卡（打卡尚未實作）
2. 選擇 Excel 檔案、輸入工作表名稱
3. 開始上傳並處理 → GAS 解析並寫入
4. 前端顯示摘要（新增筆數、略過重複）與明細卡片（含重複標示）
5. 錯誤/警告時捲動至訊息處；成功時捲動至結果區

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
├── frontend/                   # 前端網頁
│   ├── index.html              # 入口：選擇測試/正式環境
│   ├── test/                   # 測試環境前端（開發用）
│   │   ├── index.html          # 查詢班表 + 上傳（含分店下拉）
│   │   ├── config.js           # ENVIRONMENT: test, GAS_URL_TEST, TARGET_GOOGLE_SHEET_TAB: 班表
│   │   ├── test-connection.html
│   │   └── scripts/upload.js
│   └── prod/                   # 正式環境前端（deploy-prod 同步）
│       ├── index.html
│       ├── config.js           # ENVIRONMENT: production, GAS_URL_PROD, TARGET_GOOGLE_SHEET_TAB: 班表
│       ├── test-connection.html
│       └── scripts/upload.js
│
├── gas-backend/                # Google Apps Script 後端
│   ├── test/                   # 測試環境 GAS
│   │   ├── .clasp.json         # 測試環境腳本 ID（不提交）
│   │   ├── .claspignore
│   │   ├── appsscript.json
│   │   ├── Code.gs             # 主程式（doGet/doPost）
│   │   ├── Config.gs           # 環境設定、Log 系統
│   │   ├── ExcelParser.gs      # 泉威國安班表解析器
│   │   ├── CompareService.gs   # 班表與打卡比對、校正讀寫
│   │   ├── SheetService.gs     # 寫入、讀取、去重、loadSchedule、loadAttendance
│   │   ├── DiagnosticUtils.gs  # 診斷工具
│   │   └── TestUtils.gs        # 測試工具函數
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
│   └── 2026泉威國安班表.xlsx   # 實際班表檔案（11501, 11502 工作表）
│
├── docs/                       # 文件
│   ├── API.md                  # API 端點說明
│   ├── FIELD_MAPPING.md        # 欄位映射說明
│   ├── LOG_SYSTEM.md           # Log 系統說明
│   └── NO_CORS_ISSUE.md        # CORS 問題說明
│
├── REPLACEMENT_SUMMARY.md      # 解析器替換完成報告
├── TEST_REPORT.md              # 測試報告模板
├── HOW_TO_TEST.md              # 測試指南
├── test-parser.js              # Node.js 測試腳本（含 xlsx 套件）
├── verify-parser.js            # 快速驗證腳本
├── parse-result.json           # 最新解析結果（JSON）
├── parse-result.csv            # 最新解析結果（CSV）
│
└── package.json                # npm 相關（Jest、xlsx 等）
```

---

## 3. 資料流（重要）

### 3.1 泉威國安班表上傳與處理流程

```
使用者 → 上傳 Excel（泉威國安班表）+ 選擇工作表（如 11501）
       ↓
GAS doPost → /upload
       ↓
1. 檔案驗證（格式、大小）
2. Excel 轉換為 Google Sheets（透過 Drive API）
3. 讀取指定工作表的二維陣列資料
       ↓
4. 【泉威班表解析】parseQuanWeiSchedule()
   ├─ buildShiftMap() → 掃描 * 開頭的代碼定義（A, B, O...）
   ├─ locateDataStructure() → 自動定位日期列（1, 2, 3...）
   ├─ 取得年月份（從 A2 儲存格，如 2026/01）
   └─ 主解析迴圈
       ├─ 遍歷員工列
       ├─ 遍歷日期欄
       ├─ 查詢代碼字典或解析手寫時間
       ├─ 計算工作時數
       └─ 產生記錄 [員工, 日期, 上班, 下班, 時數]
       ↓
5. 轉換後的資料寫入「國安班表」工作表
6. 格式化標題列（藍底白字）
7. 建立處理記錄
8. 記錄 Log（OPERATION、DEBUG 等級）
       ↓
前端顯示成功訊息 + 處理統計
```

### 3.2 Excel 檔案格式（泉威國安班表）

#### 特殊格式說明
```
A1: 空白          B1: "更新日"
A2: "2026/01"     B2: 更新日期
A3: "姓名/星期"   B3: "入職日"   C3-Z3: 1, 2, 3, 4, 5... (橫向日期)
A4: 空白          B4-Z4: 週一, 週二, 週三...
A5: "TiNg"        B5: 入職日     C5-Z5: A, B, O, B1... (班別代碼)
A6: "茶葉"        B6: 入職日     C6-Z6: A, B, O...
...

代碼定義區（表格下方）：
A25: "* A 10:00-17:00"   (早班)
A26: "* A1 10:00-15:00"  (短早班)
A27: "* B 16:30-20:30"   (晚班)
A28: "* B1 14:30-20:30"  (長晚班)
A29: "* C 12:30-16:30"   (中班，僅 11502)
A30: "* O 10:00-20:30"   (全日班)
```

#### 必要元素
- ✅ 年月份（A2，格式：YYYY/MM）
- ✅ 日期列（包含數字 1, 2, 3...）
- ✅ 員工姓名欄（A 欄）
- ✅ 代碼定義區（* 開頭）

---

## 4. API 端點設計

所有端點透過 GAS Web App 部署，使用 `doGet` 和 `doPost` 處理。

| 方法 | action / 說明 | 參數 |
|------|---------------|------|
| GET | `test` | 連線測試 |
| GET | `getBranches` | 取得分店清單（來源「分店」sheet） |
| GET | `getPersonnelByBranch` | 依分店取得人員名單（班表名稱、打卡名稱聯集） |
| GET | `loadSchedule` | `yearMonth`(YYYYMM)、`date`(YYYY-MM-DD)、`names`(逗號分隔)、`branch`，AND 篩選 |
| GET | `loadAttendance` | `yearMonth`(YYYYMM)、`date`(YYYY-MM-DD)、`names`(逗號分隔)、`branch`，AND 篩選 |
| GET | `loadCompare` | `yearMonth`(YYYYMM) 或 `startDate`+`endDate`、`names`(逗號分隔)、`branch`(必填)；回傳班表與打卡一對一比對 items |
| POST | `upload` | `uploadType`(schedule/attendance)、`fileName`、`fileData`(base64)、`targetSheetName`、`targetGoogleSheetTab`、`branchName`(班表必填) |
| POST | `submitCorrection` | `branch`、`empAccount`、`name`、`date`、`scheduleStart`、`scheduleEnd`、`scheduleHours`、`attendanceStart`、`attendanceEnd`、`attendanceHours`、`attendanceStatus`、`correctedStart`、`correctedEnd` |
| POST | `calculate` | 薪資計算（尚未實作） |

---

## 5. Google Sheets 表結構

專案使用一個 Google Sheets 檔案，**目前已實作**的工作表：

| 工作表名稱 | 說明 | 主要欄位 |
|------------|------|----------|
| **班表** | 上傳的班表解析結果、loadSchedule 查詢來源 | 員工姓名、排班日期、上班時間、下班時間、工作時數、班別、分店 |
| **打卡** | 上傳的打卡 CSV 結果 | 分店、員工編號、員工帳號、員工姓名、打卡日期、上班、下班、時數、狀態 |
| **人員** | 班表名稱／打卡名稱 ↔ 員工帳號對應 | 員工帳號、班表名稱、打卡名稱、分店 |
| **校正** | 班表與打卡校正紀錄 | 分店、員工帳號、姓名、日期、班表上下班、打卡上下班、校正上下班、是否有效、建立時間 |
| **Log** | 系統 Log | 時間、等級、訊息、環境、詳細資訊 |
| **處理記錄** | 上傳處理記錄 | 時間、檔案名稱、來源/目標工作表、筆數、狀態、訊息 |

*其他工作表（計算結果、調整記錄等）為未來薪資計算功能預留。*

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
- **部署 URL**：`https://script.google.com/macros/s/AKfycbxpExUCOrTLcbUSHcaBmtluODqb2pNopEuh7QOUJXJXXVZcKMP_EJiUKRLsTTZ3sANS5Q/exec`
- **Script Properties**：
  - `SHEET_ID`: 測試用 Google Sheets ID
  - `ENVIRONMENT`: `test`
  - `Log_Level`: `2`（建議，Debug 等級）
- **前端設定**：`frontend/test/config.js` 中 `GAS_URL_TEST`

### 8.2 正式環境
- **GAS 專案**：SalarySystem - Production
- **Script ID**：`185v2ySyq9ncyzBwbnnogQTXHuw3S78pfmXcaATxWRs5K5W3ZYsSRG4F7`
- **Script Properties**：
  - `SHEET_ID`: 正式用 Google Sheets ID
  - `ENVIRONMENT`: `production`
  - `Log_Level`: `1`（建議，營運等級，只記錄重要操作與錯誤）
- **前端設定**：`frontend/prod/config.js` 中 `GAS_URL_PROD`

### 8.3 開發流程

#### 日常開發（使用 `go` 指令）
1. 在 `gas-backend/test/` 開發並測試
2. 執行 `go` 指令：
   - 自動 commit & push GitHub
   - 自動 `clasp push` 到測試環境
   - 自動更新 CONTEXT.md
3. 使用 `frontend/test/index.html` 驗證功能

#### 部署到正式環境（需明確指令）
1. 測試環境驗證完成後，**明確告知 AI** 要部署到正式環境
2. AI 執行 `deploy-prod`：
   - 複製 `gas-backend/test/*.gs` 到 `gas-backend/prod/`（排除 `.clasp.json`）
   - 複製 `frontend/test/*` 到 `frontend/prod/`（config.js 保留 prod 專用設定）
   - `cd gas-backend/prod && clasp push`
   - Commit 並 push 到 GitHub
   - 更新部署記錄

---

## 9. 開發與部署

### 9.1 GitHub
- **遠端**：`origin` → https://github.com/chyuanwei/SalarySystem.git
- **分支**：`master`

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
| **go（限制）** | **僅能操作測試環境（test）**，**不得更改正式環境（prod）程式碼**；正式環境任何變更需使用者明確指示（如 `deploy-prod`）。 |
| **deploy-prod** | 部署到正式環境：複製 `gas-backend/test/*.gs` 到 `gas-backend/prod/` → 複製 `frontend/test/*` 到 `frontend/prod/`（保留 prod config.js）→ `cd gas-backend/prod && clasp push` → commit/push GitHub **（需附 commit message）** → 更新部署記錄 → 提供 **Comment for GAS deploy**。**僅在使用者明確指示時執行**。 |
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

## 12. 已實作功能

### 12.1 查詢國安班表
- ✅ 日期篩選：整月（年+月 或 YYYYMM）、單日
- ✅ 人員多選（載入後取得名單，AND 篩選）
- ✅ loadSchedule API（yearMonth、date、names）
- ✅ 結果卡片與摘要
- ✅ 上班/下班時間正確顯示（normalizeTimeValue 支援 Date/數字）

### 12.2 檔案上傳與解析
- ✅ 上傳類型單選：班表 | 打卡（打卡尚未實作）
- ✅ 拖曳或選擇 Excel 檔案上傳
- ✅ 檔案大小與格式驗證、上傳進度顯示
- ✅ 可自訂 Excel 工作表名稱（11501、11502 等）
- ✅ RWD 響應式設計

### 12.3 泉威國安班表專用解析器
- ✅ 支援 .xlsx 和 .xls 格式
- ✅ 自動轉換 Excel 為 Google Sheets（透過 Drive API）
- ✅ **自動掃描班別代碼定義**（* A 10:00-17:00 格式）
- ✅ **自動定位資料結構**（搜尋日期列 1, 2, 3...）
- ✅ **支援代碼式班別**（A, A1, B, B1, C, O 等）
- ✅ **支援手寫時間格式**（如 18:00-20:30）
- ✅ **自動計算工作時數**（含跨日處理）
- ✅ **橫向日期佈局支援**
- ✅ **年月份自動偵測**（從 A2 讀取）

### 12.4 核心解析函數
- ✅ `parseQuanWeiSchedule()` - 主解析函數
- ✅ `buildShiftMap()` - 代碼字典掃描
- ✅ `locateDataStructure()` - 自動定位資料結構
- ✅ `parseTimeRange()` - 時間範圍解析與時數計算
- ✅ `formatTimeString()` - 時間格式化（HH:MM）
- ✅ `formatScheduleDate()` - 日期格式化（YYYY/MM/DD）

### 12.5 資料儲存
- ✅ 寫入「國安班表」Google Sheets 工作表
- ✅ 標準格式：員工姓名 | 排班日期 | 上班時間 | 下班時間 | 工作時數 | 班別（第 F 欄）
- ✅ 自動格式化標題列（藍底白字、粗體）
- ✅ 自動調整欄寬
- ✅ 凍結標題列
- ✅ 附加模式寫入（不覆蓋現有資料）
- ✅ 寫入前去重（同一筆資料不重複追加）
- ✅ 去重欄位標準化（日期/時間統一格式，避免重複寫入）

### 12.6 前端 UX
- ✅ 上傳結果：摘要（新增筆數、略過重複、原始筆數）、卡片列表
- ✅ 重複標記：每筆 isDuplicate、卡片上方置中「重複」標籤
- ✅ 錯誤/警告時捲動至訊息處；成功時捲動至結果區
- ✅ 連線測試逾時 15 秒

### 12.7 Log 系統
- ✅ 完整的 Log 架構（DEBUG / INFO / WARNING / ERROR / OPERATION）
- ✅ 透過 Script Properties `Log_Level` 控制 Log 等級（0=關閉, 1=營運, 2=Debug）
- ✅ 記錄到 Google Sheets 的 `Log` 工作表
- ✅ 包含時間、等級、訊息、環境、使用者、詳細資訊
- ✅ 便利函數：logDebug(), logInfo(), logWarning(), logError(), logOperation()
- ✅ 自動清理舊 Log 功能（cleanOldLogs）
- ✅ 詳細的 Log 說明文件（docs/LOG_SYSTEM.md）
- ✅ **解析過程詳細記錄**（代碼掃描、資料結構、員工統計）

### 12.8 測試工具
- ✅ `TestUtils.gs` - GAS 環境測試函數
- ✅ `test-parser.js` - Node.js 完整測試腳本（含 xlsx）
- ✅ `verify-parser.js` - 快速驗證腳本
- ✅ 測試指南文件（HOW_TO_TEST.md）

### 12.9 班表與打卡比對
- ✅ 條件選擇：月份(YYYYMM) 或 日期區間、分店(必選)、人員(可複選)
- ✅ loadCompare API：一對一比對班表與打卡，依人員 sheet 對應班表名稱/打卡名稱
- ✅ 比對結果卡片：班表 vs 打卡（分店、員工帳號、姓名、日期、上班/下班/時數/狀態）
- ✅ 校正：校正上班/下班輸入、校正送出、已校正顯示唯讀+編輯按鈕
- ✅ submitCorrection API：寫入校正 sheet，再次校正時舊筆設為無效
- ✅ getPersonnelByBranch：選擇分店時載入該分店人員（方案 C：載入比對後合併結果人員）
- ✅ 比對卡片：手機優先排版（姓名+日期主標題、班表/打卡分塊、校正區垂直排列、44px 觸控目標）
- ✅ CompareService：getPersonnelByBranch、readPersonnelMapping、readScheduleByConditions、readAttendanceByConditions、compareScheduleAttendance、readCorrectionsValid、writeCorrection

### 12.10 測試與正式環境
- ✅ 完整的環境分離（GAS + 前端）
- ✅ 獨立的 GAS Script ID（gas-backend/test/、gas-backend/prod/）
- ✅ 前端 test/prod 分離（frontend/test/、frontend/prod/，入口頁選擇環境）
- ✅ 獨立的 Google Sheets

---

## 13. 待實作功能（未來擴充）

### 13.1 打卡上傳
- ✅ 打卡 CSV 上傳：人員驗證、分店 mapping、去重、寫入「打卡」sheet

### 13.2 多格式支援
- ☐ 支援其他班表格式
- ☐ 自動偵測班表類型

### 13.3 進階功能
- ☐ 支援人工調整班次、調整理由記錄
- ☐ 歷史版本對比

### 13.4 報表匯出
- ☐ 匯出為 Excel 格式
- ☐ 匯出為 PDF 格式
- ☐ 統計報表（員工工時統計、班別分布等）

### 13.5 薪資計算（視需求）
- ☐ 根據工時計算薪資
- ☐ 套用不同費率規則
- ☐ 處理加班費計算

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

## 15. 附錄：曾遇問題與注意

| 狀況 | 說明／解法 |
|------|------------|
| **Git index.lock** | 若 `git status` 報 index.lock 錯誤，多為上次操作中斷；可關閉其他 Git 程式後刪除 `.git/index.lock`。 |
| **clasp push 失敗（rootDir）** | 錯誤提示 rootDir 不匹配時，編輯 `.clasp.json`，將 `rootDir` 改為實際目錄路徑。 |
| **測試與正式環境混淆** | 務必確認當前操作的是 `gas-backend/test/` 還是 `gas-backend/prod/`，避免誤推。 |
| **Excel 解析錯誤** | 確認 Excel 檔案格式為 .xlsx 或 .xls，且欄位名稱與預期一致。 |
| **GAS 執行逾時** | GAS 單次執行限制 6 分鐘，若資料量大需考慮分批處理。 |

---

## 16. 重要注意事項

### 16.1 CORS 與前端結果顯示

前端已改為 **可讀取回應（CORS 模式）**，上傳後會直接顯示解析結果。

**若出現 CORS 錯誤**：
1. 確認 GAS 已重新部署為 Web App
2. 確認回應有 `Access-Control-Allow-Origin`（由 `createJsonResponse` 設定）
3. 若仍失敗，暫時用 `frontend/test/test-connection.html` 驗證 GAS 連線

**備註**：若 CORS 被瀏覽器阻擋，前端將無法顯示結果，但後端仍可能已完成寫入。

---

## 17. 部署記錄

| 日期 | 版本 | 環境 | 更新內容 | 部署者 |
|------|------|------|----------|--------|
| 2026-02-07 | v0.6.27 | 測試 | 校正 POST 改 Content-Type text/plain 避 CORS；打卡區塊格式（時間+時數+狀態、無狀態則省略）；formatHours 避免 7小時h | AI |
| 2026-02-07 | v0.6.26 | 測試 | getPersonnelByBranch（選分店載入人員）；比對卡片手機優先排版 | AI |
| 2026-02-07 | v0.6.25 | 測試 | 班表與打卡比對：loadCompare、submitCorrection、CompareService；前端比對區塊、校正卡片、已校正/編輯狀態 | AI |
| 2026-02-07 | v0.6.24 | 測試 | 查詢打卡功能（loadAttendance）；打卡 sheet 欄位順序調整為 分店,員工編號,員工帳號,員工姓名,... | AI |
| 2026-02-07 | v0.6.23 | 測試 | 修正 appendAttendance getRange 列數不符；檔案輸入 accept 加入 .csv。 | AI |
| 2026-02-06 | v0.6.22 | 測試 | 打卡 CSV 上傳：人員驗證、分店 mapping、去重、寫入「打卡」sheet；GitHub Actions 部署 Pages。 | AI |
| 2026-02-06 | v0.6.21 | 測試 | 修正查詢區分店下拉無資料：loadBranches 改為執行時查詢 DOM，確保上傳與查詢區皆正確填入。 | AI |
| 2026-02-06 | v0.6.20 | 測試 | UI 國安班表→班表；查詢區新增分店下拉篩選；loadSchedule 支援 branch 參數；readFromSheet 不建空白 sheet。 | AI |
| 2026-02-06 | v0.6.19 | 測試 | 班表新增欄位 G「分店」；工作表改名為「班表」；上傳畫面分店下拉（來源「分店」sheet）；getBranches API；buildDedupKey 含分店。 | AI |
| 2026-02-06 | v0.6.18 | 測試 | 前端 test/prod 分離：入口 index.html、frontend/test/、frontend/prod/；開發僅改 test，deploy-prod 同步至 prod。 | AI |
| 2026-02-06 | v0.6.17 | 正式 | 從 test 同步至 prod：loadSchedule、uploadType、allRecordsWithFlag；Code/Config 設為正式環境。 | AI |
| 2026-02-06 | v0.6.17 | 測試 | 修正工作表名稱為空時未捲動至錯誤；查詢國安班表區塊 UX 重新排版。 | User |
| 2026-02-06 | v0.6.16 | 測試 | 錯誤/警告時捲動至訊息處；成功時捲動至結果區。 | User |
| 2026-02-06 | v0.6.15 | 測試 | 上傳新增「班表」/「打卡」單選；打卡功能尚未實作。 | User |
| 2026-02-06 | v0.6.14 | 測試 | 修正國安班表查詢的上班/下班時間顯示；normalizeTimeValue 支援數字（日的小數）格式。 | User |
| 2026-02-06 | v0.6.13 | 測試 | 新增國安班表查詢：YYYYMM/日期選擇器 + 人員多選，AND 篩選；loadSchedule API 支援 date、names 參數。 | User |
| 2026-02-06 | v0.6.12 | 測試 | 前端連線測試逾時改為 15 秒；結果卡片「重複」標籤改為上方置中顯示。 | User |
| 2026-02-06 | v0.6.11 | 測試 | 上傳結果摘要新增「略過重複」筆數；每筆結果回傳 `isDuplicate` 並在前端卡片標示「重複」。 | User |
| 2026-02-06 | v0.6.10 | 測試 | 去重比對改為標準化日期/時間格式，避免重複資料誤判。 | User |
| 2026-02-06 | v0.6.9 | 測試 | 國安班表寫入前去重（姓名+日期+上下班+班別），回傳實際新增筆數與略過數量。 | User |
| 2026-02-06 | v0.6.8 | 測試 | 前端顯示解析結果（摘要+卡片列表，RWD 優化）；後端回傳 JSON 結果並加 CORS 標頭；前端改用可讀取回應模式。 | User |
| 2026-02-06 | v0.6.7 | 測試 | 國安班表新增第 F 欄「班別」；修正 buildShiftMap 動態掃描與終止條件（空列只 skip）；Utilities.newBlob 與 Log 權限修正；新增 DiagnosticUtils.gs。 | User |
| 2026-02-06 | v0.6 | 測試+正式 | **重大更新**：替換為泉威國安班表專用解析器。新增：自動掃描代碼定義、自動定位資料結構、代碼式班別解析、橫向日期支援、時數自動計算。移除：通用欄位映射邏輯。新增測試工具（TestUtils.gs）和驗證腳本。 | AI |
| 2026-02-06 | v0.5 | 測試+正式 | 實作自動欄位映射與格式轉換：支援多種欄位名稱變體、自動計算工作時數、日期時間格式化、附加模式寫入 | User |
| 2026-02-06 | v0.4 | 測試 | 新增測試工具函數（TestUtils.gs），用於分析 Google Sheets 結構與格式，更新預設 SHEET_ID | User |
| 2026-02-06 | v0.3 | 測試 | 更新 GAS 部署 URL、改善錯誤處理、新增連線測試頁面、No-CORS 問題說明 | User |
| 2026-02-06 | v0.2 | 測試+正式 | 新增完整 Log 系統，支援 Log_Level 控制（0/1/2），五種 Log 等級，記錄到 'Log' sheet | AI |
| 2026-02-06 | v0.1 | 測試+正式 | 初始版本：Excel 上傳、解析指定工作表、寫入 Google Sheets、基本 Log 功能 | AI |

---

## 18. 泉威國安班表解析邏輯

### 目標格式（國安班表）
- **欄位**：員工姓名 | 排班日期 | 上班時間 | 下班時間 | 工作時數 | 班別（F 欄）
- **日期格式**：YYYY/MM/DD（例如：2026/01/15）
- **時間格式**：HH:MM（例如：08:00）
- **工作時數**：自動計算（單位：小時，精確到小數點後 1 位）

### 解析流程（parseQuanWeiSchedule）

#### 步驟 1: 掃描班別代碼定義
```javascript
buildShiftMap(data)
// 搜尋 * 開頭的列，格式：* A 10:00-17:00
// 提取代碼（A）和時間範圍（10:00-17:00）
// 自動計算工作時數（含跨日處理）
```

#### 步驟 2: 定位資料結構
```javascript
locateDataStructure(data)
// 搜尋前 10 列中包含數字 1 的儲存格
// 定位日期列、日期欄、員工起始列
```

#### 步驟 3: 取得年月份
```javascript
// 從 A2 讀取年月份（格式：2026/01）
// 若無法取得則使用預設值
```

#### 步驟 4: 主解析迴圈
```javascript
// 遍歷員工列（從 startRow 開始）
for (員工) {
  // 終止條件：空白姓名或統計關鍵字
  
  // 遍歷日期欄
  for (日期) {
    // 取得儲存格值（代碼或手寫時間）
    
    // 查詢代碼字典或解析手寫時間
    if (shiftMap[代碼]) {
      // 使用代碼定義
    } else if (cellValue.includes('-')) {
      // 解析手寫時間
    }
    
    // 產生記錄 [員工, 日期, 上班, 下班, 時數]
  }
}
```

### 支援的班別代碼
- **A** - 早班（例如：10:00-17:00）
- **A1** - 短早班（例如：10:00-15:00）
- **B** - 晚班（例如：16:30-20:30）
- **B1** - 長晚班（例如：14:30-20:30）
- **C** - 中班（例如：12:30-16:30，僅 11502）
- **O** - 全日班（例如：10:00-20:30）
- **手寫時間** - 支援 HH:MM-HH:MM 格式

### 終止條件
解析會在遇到以下情況時停止：
- 空白員工姓名列
- 包含關鍵字：上班人數、合計、備註、P.T、閉店評論

### 資料寫入模式
- **附加模式**：資料會附加到「國安班表」工作表末尾
- **第一次上傳**：包含標題列並格式化
- **後續上傳**：只附加資料列

**詳細說明**：見 `REPLACEMENT_SUMMARY.md` 和 `HOW_TO_TEST.md`

---

## 19. v0.6 版本重大更新說明

### 📌 重大變更：解析器完全替換

#### 變更原因
原系統設計為通用班表解析器，但實際需求為處理「泉威國安班表」的特定格式：
- 橫向日期佈局（1, 2, 3...）
- 代碼式班別（A, B, O...）
- 代碼定義區（* A 10:00-17:00）

通用解析器無法有效處理此格式，因此進行專用化改造。

#### 主要變更內容

**移除的功能（~270 行）：**
- ❌ `transformToTargetFormat()` - 通用格式轉換
- ❌ `detectColumnMapping()` - 通用欄位映射
- ❌ `calculateWorkHours()` - 通用時數計算
- ❌ 其他通用欄位處理函數

**新增的功能（~240 行）：**
- ✅ `parseQuanWeiSchedule()` - 泉威班表主解析器
- ✅ `buildShiftMap()` - 代碼字典掃描
- ✅ `locateDataStructure()` - 自動定位資料結構
- ✅ `parseTimeRange()` - 時間範圍解析
- ✅ `formatTimeString()` - 時間格式化
- ✅ `formatScheduleDate()` - 日期格式化

**測試工具：**
- ✅ `TestUtils.gs` - GAS 環境測試函數
- ✅ `test-parser.js` - Node.js 完整測試（含 xlsx）
- ✅ `verify-parser.js` - 快速本地驗證
- ✅ `HOW_TO_TEST.md` - 詳細測試指南
- ✅ `REPLACEMENT_SUMMARY.md` - 替換完成報告

#### 驗證狀態

**✅ 本地驗證完成：**
- 代碼掃描：找到 5 個代碼（A, A1, B, B1, O）
- 資料結構定位：成功
- 員工排班解析：3 位員工，9 筆記錄
- 時間與日期格式化：正確

**⏳ 待執行測試：**
- [ ] GAS 編輯器內測試（`testParseQuanWeiSchedule()`）
- [ ] 實際檔案上傳測試（11501, 11502）
- [ ] Google Sheets 輸出驗證
- [ ] Log 記錄檢查

#### 測試方法

**快速驗證（本地）：**
```bash
node verify-parser.js
```

**GAS 環境測試：**
```bash
cd gas-backend/test
npx clasp open
# 執行 testParseQuanWeiSchedule()
```

**實際上傳測試：**
1. 開啟 `frontend/index.html`（入口）→ 選擇測試環境 → `frontend/test/index.html`
2. 上傳 `2026泉威國安班表.xlsx`
3. 選擇工作表 `11501` 或 `11502`
4. 檢查「國安班表」和「Log」工作表

#### 重要注意事項

**⚠️ 格式限制：**
- 僅支援泉威國安班表格式
- 不再支援通用班表格式
- 代碼定義必須以 `*` 開頭
- 日期列必須包含數字 1
- 年月份必須在 A2（格式：YYYY/MM）

**⚠️ 終止條件：**
- 空白員工姓名列
- 關鍵字：上班人數、合計、備註、P.T、閉店評論

#### 相關文件
- 📄 `REPLACEMENT_SUMMARY.md` - 完整替換報告
- 📄 `HOW_TO_TEST.md` - 測試指南（⭐ 必讀）
- 📄 `TEST_REPORT.md` - 測試報告模板
- 📄 `verify-parser.js` - 快速驗證腳本
- 📄 `test-parser.js` - 完整測試腳本

---

*本檔案為專案專用 context，請隨重要變更更新。*
*最後更新：2026-02-07（v0.6.27 CORS fix、打卡格式、formatHours）*
