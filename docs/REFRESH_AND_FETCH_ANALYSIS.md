# Refresh 與 Fetch 行為分析

## 問題 1：為何頁面 refresh 就會跑到「上傳」的 tab？

### 原因
`frontend/test/index.html` 中，**上傳 tab 被寫死為預設 active**：

```html
<!-- 上傳按鈕：有 active class -->
<button type="button" class="tab-nav-btn active" ... id="tab-btn-upload" data-tab="upload">📄 上傳</button>
<!-- 查詢、比對、計算：無 active -->
<button type="button" class="tab-nav-btn" ... data-tab="query">📅 查詢</button>
...

<!-- 上傳內容區：有 active -->
<div id="tab-pane-upload" class="tab-pane active" ...>
<!-- 查詢、比對：有 hidden -->
<div id="tab-pane-query" class="tab-pane" ... hidden>
```

頁面載入時 HTML 直接指定「上傳」為 active、其它為 hidden，因此每次 refresh 都會回到上傳 tab。

### 若要改為預設顯示查詢
將 `active` 與 `hidden` 從上傳移到查詢：
- `tab-btn-upload`：移除 `active`，`aria-selected="false"`
- `tab-btn-query`：加上 `active`，`aria-selected="true"`
- `tab-pane-upload`：移除 `active`，加上 `hidden`
- `tab-pane-query`：加上 `active`，移除 `hidden`

並調整 `initTabNav()` 內 `updateSharedConditionVisibility()` 的初始值為 `'query'`。

---

## 問題 2：為何每次 refresh 都會做 fetch？

### 載入流程
`upload.js` 在 `DOMContentLoaded` 時執行：

```javascript
document.addEventListener('DOMContentLoaded', function() {
  initTabNav();
  initYearMonthSelects();
  toggleDateFilterMode();
  loadBranches();  // ← 每次載入都會呼叫
  // ...
});
```

### `loadBranches()` 會做什麼
1. **fetch getBranches**：呼叫 GAS `?action=getBranches`，取得分店清單
2. 成功後：更新 `branchSelect`、`queryBranchSelect`、`compareBranchSelect` 的 options
3. 接著呼叫 **`loadQueryPersonnel()`**

### `loadQueryPersonnel()` 的行為
- 需要：`branch` 已選、以及 `yearMonth` 或 `startDate`
- 初始化時分店未選、年月未選，會 early return，**不發送 fetch**
- 因此初始化時實際發出的請求只有：**1 次 getBranches**

### 結論
- 每次 refresh 至少會有一次 fetch：`getBranches`
- 原因：`loadBranches()` 在 `DOMContentLoaded` 被呼叫，用於載入分店下拉，供上傳／查詢／比對共用

### 若要減少或延遲 fetch
1. **快取分店**：將 getBranches 結果存 `sessionStorage`，有快取時不 fetch
2. **延遲載入**：等到使用者切到需要分店的 tab 時再呼叫 `loadBranches()`
3. **保留現狀**：分店清單可能變動，每次載入更新是較保守做法
