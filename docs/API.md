# API 端點說明

本文件說明薪資計算系統的 API 端點設計。

所有 API 端點透過 Google Apps Script Web App 提供，使用 `doGet` 和 `doPost` 方法處理請求。

---

## 基本資訊

- **Base URL（測試環境）**: `https://script.google.com/macros/s/{TEST_SCRIPT_ID}/exec`
- **Base URL（正式環境）**: `https://script.google.com/macros/s/{PROD_SCRIPT_ID}/exec`
- **Content-Type**: `application/json`
- **編碼**: UTF-8

---

## API 端點列表

### 1. 上傳檔案

上傳班表與打卡紀錄 Excel 檔案。

- **端點**: `/upload`
- **方法**: `POST`
- **參數**:

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| `scheduleFile` | String | 是 | 班表 Excel 檔案（Base64 編碼） |
| `attendanceFile` | String | 是 | 打卡紀錄 Excel 檔案（Base64 編碼） |
| `fileName1` | String | 是 | 班表檔案名稱 |
| `fileName2` | String | 是 | 打卡紀錄檔案名稱 |

- **請求範例**:

```javascript
{
  "action": "upload",
  "scheduleFile": "base64EncodedString...",
  "attendanceFile": "base64EncodedString...",
  "fileName1": "班表202401.xlsx",
  "fileName2": "打卡紀錄202401.xlsx"
}
```

- **回應範例（成功）**:

```javascript
{
  "success": true,
  "jobId": "job_20240101_123456",
  "message": "檔案已上傳，開始處理"
}
```

- **回應範例（失敗）**:

```javascript
{
  "success": false,
  "error": "檔案格式錯誤",
  "details": "scheduleFile 不是有效的 Excel 檔案"
}
```

---

### 2. 查詢處理狀態

查詢特定 job 的處理狀態。

- **端點**: `/status`
- **方法**: `GET`
- **參數**:

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| `jobId` | String | 是 | Job ID |

- **請求範例**:

```
GET /status?jobId=job_20240101_123456
```

- **回應範例（處理中）**:

```javascript
{
  "success": true,
  "jobId": "job_20240101_123456",
  "status": "processing",
  "message": "正在計算薪資...",
  "progress": 50
}
```

- **回應範例（完成）**:

```javascript
{
  "success": true,
  "jobId": "job_20240101_123456",
  "status": "completed",
  "message": "計算完成",
  "completedAt": "2024-01-01T12:35:00Z"
}
```

- **狀態說明**:
  - `uploading`: 上傳中
  - `processing`: 處理中
  - `completed`: 完成
  - `error`: 錯誤

---

### 3. 取得計算結果

取得特定 job 的計算結果。

- **端點**: `/result`
- **方法**: `GET`
- **參數**:

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| `jobId` | String | 是 | Job ID |

- **請求範例**:

```
GET /result?jobId=job_20240101_123456
```

- **回應範例**:

```javascript
{
  "success": true,
  "jobId": "job_20240101_123456",
  "data": [
    {
      "employeeId": "E001",
      "employeeName": "張三",
      "workDays": 22,
      "totalHours": 176,
      "normalHours": 176,
      "overtimeHours": 0,
      "salary": {
        "base": 35200,
        "overtime": 0,
        "total": 35200
      },
      "attendance": {
        "late": 0,
        "earlyLeave": 0,
        "absent": 0
      },
      "anomalies": []
    },
    {
      "employeeId": "E002",
      "employeeName": "李四",
      "workDays": 20,
      "totalHours": 170,
      "normalHours": 160,
      "overtimeHours": 10,
      "salary": {
        "base": 32000,
        "overtime": 2680,
        "total": 34680
      },
      "attendance": {
        "late": 2,
        "earlyLeave": 0,
        "absent": 2
      },
      "anomalies": [
        {
          "date": "2024-01-05",
          "type": "missing_punch",
          "message": "缺少下班打卡"
        }
      ]
    }
  ]
}
```

---

### 4. 提交調整資料

使用者檢視結果後，提交調整資料。

- **端點**: `/adjust`
- **方法**: `POST`
- **參數**:

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| `jobId` | String | 是 | Job ID |
| `adjustments` | Array | 是 | 調整項目陣列 |

- **請求範例**:

```javascript
{
  "action": "adjust",
  "jobId": "job_20240101_123456",
  "adjustments": [
    {
      "employeeId": "E002",
      "field": "overtimeHours",
      "oldValue": 10,
      "newValue": 8,
      "reason": "實際加班時數應為 8 小時，原計算誤判"
    },
    {
      "employeeId": "E003",
      "field": "salary.total",
      "oldValue": 30000,
      "newValue": 32000,
      "reason": "補發全勤獎金 2000 元"
    }
  ]
}
```

- **回應範例（成功）**:

```javascript
{
  "success": true,
  "jobId": "job_20240101_123456",
  "message": "調整已儲存",
  "adjustedCount": 2
}
```

---

### 5. 確認最終結果

使用者確認調整後的結果（二次確認）。

- **端點**: `/confirm`
- **方法**: `POST`
- **參數**:

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| `jobId` | String | 是 | Job ID |

- **請求範例**:

```javascript
{
  "action": "confirm",
  "jobId": "job_20240101_123456"
}
```

- **回應範例（成功）**:

```javascript
{
  "success": true,
  "jobId": "job_20240101_123456",
  "message": "已確認並儲存最終結果",
  "confirmedAt": "2024-01-01T13:00:00Z"
}
```

---

### 6. 匯出報表

匯出特定 job 的薪資報表。

- **端點**: `/export`
- **方法**: `GET`
- **參數**:

| 參數名稱 | 類型 | 必填 | 說明 |
|---------|------|------|------|
| `jobId` | String | 是 | Job ID |
| `format` | String | 否 | 匯出格式（`xlsx` 或 `pdf`，預設 `xlsx`） |

- **請求範例**:

```
GET /export?jobId=job_20240101_123456&format=xlsx
```

- **回應**: 檔案下載

---

## 錯誤碼

| 錯誤碼 | 說明 |
|--------|------|
| 400 | 請求參數錯誤 |
| 404 | 找不到指定的 Job ID |
| 500 | 伺服器內部錯誤 |

---

## 錯誤回應格式

```javascript
{
  "success": false,
  "error": "錯誤訊息",
  "code": 400,
  "details": "詳細錯誤說明"
}
```

---

## 測試

可使用以下工具測試 API：

- **Postman**: 匯入 API collection（待建立）
- **cURL**: 命令列測試
- **前端測試頁**: `frontend/index.html`

---

**更新日期**: 2024-01-01
