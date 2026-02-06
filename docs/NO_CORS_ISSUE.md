# No-CORS 模式問題說明

本文件說明前端使用 `no-cors` 模式的限制與解決方案。

---

## 🔍 問題說明

### 什麼是 No-CORS？

當前前端程式碼使用 `no-cors` 模式發送請求：

```javascript
const response = await fetch(CONFIG.GAS_URL, {
  method: 'POST',
  mode: 'no-cors', // ← 這個設定
  headers: {
    'Content-Type': 'application/json'
  },
  body: JSON.stringify(payload)
});
```

### No-CORS 的限制

1. **無法讀取回應**：
   - Response 永遠是 `opaque` type
   - 無法知道請求是成功 (200) 還是失敗 (404, 500)
   - 無法讀取回應內容（JSON、錯誤訊息等）

2. **無法處理錯誤**：
   - 即使 GAS URL 錯誤，不會拋出異常
   - 即使 GAS 未部署，不會拋出異常
   - 即使處理失敗，不會拋出異常

3. **看起來總是成功**：
   - 因為無法讀取錯誤，所以程式繼續執行
   - 前端會顯示「成功」訊息
   - 但實際上可能根本沒成功

### 為什麼使用 No-CORS？

因為 GAS Web App 預設不支援 CORS（跨域資源共享），瀏覽器會阻擋跨域請求。使用 `no-cors` 可以繞過瀏覽器的 CORS 檢查，讓請求能夠發送出去。

---

## 💡 解決方案

### 方案 1：改善現有程式（已實作）

**做法**：
1. ✅ 提供連線測試頁面（`test-connection.html`）
2. ✅ 上傳前先測試連線
3. ✅ 顯示警告訊息，提醒使用者手動確認結果
4. ✅ 提供檢查清單

**優點**：
- 不需要修改 GAS
- 簡單快速

**缺點**：
- 仍然無法確定是否真的成功
- 需要使用者手動確認

---

### 方案 2：使用 JSONP（推薦）

**做法**：
1. GAS `doGet` 回傳 JSONP 格式
2. 前端使用 `<script>` 標籤載入
3. 透過 callback 函數接收結果

**GAS 程式碼**：
```javascript
function doGet(e) {
  const callback = e.parameter.callback;
  const data = { success: true, message: '測試成功' };
  
  const output = ContentService.createTextOutput(
    callback + '(' + JSON.stringify(data) + ')'
  );
  output.setMimeType(ContentService.MimeType.JAVASCRIPT);
  return output;
}
```

**前端程式碼**：
```javascript
function sendRequest(data) {
  return new Promise((resolve, reject) => {
    // 建立唯一的 callback 名稱
    const callbackName = 'jsonpCallback_' + Date.now();
    
    // 定義 callback 函數
    window[callbackName] = function(response) {
      delete window[callbackName];
      document.body.removeChild(script);
      resolve(response);
    };
    
    // 建立 script 標籤
    const script = document.createElement('script');
    script.src = CONFIG.GAS_URL + 
      '?callback=' + callbackName + 
      '&data=' + encodeURIComponent(JSON.stringify(data));
    
    script.onerror = () => {
      delete window[callbackName];
      document.body.removeChild(script);
      reject(new Error('請求失敗'));
    };
    
    document.body.appendChild(script);
  });
}
```

**優點**：
- ✅ 可以讀取回應
- ✅ 可以處理錯誤
- ✅ 不需要 CORS 設定

**缺點**：
- 只支援 GET 請求
- 需要修改 GAS 程式碼

---

### 方案 3：狀態輪詢 API

**做法**：
1. 上傳時產生 Job ID
2. 前端定期呼叫狀態 API 查詢進度
3. 根據狀態顯示結果

**流程**：
```
上傳 → 取得 Job ID → 輪詢狀態
       ↓
   處理中... (每 2 秒查詢一次)
       ↓
   完成 / 失敗
```

**優點**：
- ✅ 適合長時間處理
- ✅ 可以顯示實際進度
- ✅ 可以處理錯誤

**缺點**：
- 需要實作狀態管理
- 需要額外的 API 端點

---

### 方案 4：使用 Google Apps Script 的 HTML Service（終極方案）

**做法**：
1. GAS 回傳一個 HTML 頁面（包含表單）
2. 表單直接 submit 到 GAS
3. GAS 處理後回傳結果頁面

**優點**：
- ✅ 完全不會有 CORS 問題
- ✅ 可以正確處理錯誤

**缺點**：
- 需要重新設計前端架構
- 前端不能放在 GitHub Pages

---

## 🎯 建議

### 短期（目前實作）
- ✅ 使用 `test-connection.html` 測試連線
- ✅ 提供清楚的警告訊息
- ✅ 要求使用者手動確認結果

### 中期
- 實作**方案 2（JSONP）**或**方案 3（狀態輪詢）**
- 提供準確的成功/失敗回饋

### 長期
- 考慮使用專業的後端服務（非 GAS）
- 或使用 GAS HTML Service

---

## 🧪 如何測試

### 測試 1：使用連線測試頁面
1. 開啟 `frontend/test-connection.html`
2. 查看當前設定
3. 點擊「測試 GAS 連線」
4. 點擊「GAS 測試連結」
5. 確認是否看到 JSON 回應

### 測試 2：手動開啟 GAS URL
直接在瀏覽器開啟：
```
https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec?action=test
```

應該看到：
```json
{"success":true,"message":"測試連線成功","environment":"test"}
```

### 測試 3：檢查 Google Sheets
上傳檔案後：
1. 開啟 Google Sheets
2. 檢查是否有新的工作表
3. 檢查「Log」工作表是否有記錄

---

## 📝 目前的狀況

### 為什麼沒有報錯？

**原因**：`no-cors` 模式下，即使：
- ❌ GAS URL 錯誤
- ❌ GAS 未部署
- ❌ 處理失敗

前端仍然會：
- ✅ 顯示「成功」訊息
- ✅ 進度條跑到 100%
- ✅ 不會拋出任何錯誤

**唯一能確定的方式**：手動檢查 Google Sheets

---

## 🔧 目前的改善

**已實作**：
1. ✅ 新增連線測試頁面
2. ✅ 上傳前測試連線（雖然不太準確）
3. ✅ 顯示警告訊息
4. ✅ 提示使用者手動確認
5. ✅ 提供檢查清單

**下一步**：
- 實作 JSONP 或狀態輪詢 API
- 提供準確的錯誤處理

---

**更新日期**: 2024-02-06
