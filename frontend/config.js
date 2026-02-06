/**
 * 薪資計算系統 - 環境設定
 */

const CONFIG = {
  // GAS Web App URL
  GAS_URL_TEST: 'https://script.google.com/macros/s/1b_UcoqPQ8Vm6lekQli79JRePKDB-5qbbkLPh5lsfOdZOVtAD-Z61X8ZH/exec',
  GAS_URL_PROD: 'https://script.google.com/macros/s/185v2ySyq9ncyzBwbnnogQTXHuw3S78pfmXcaATxWRs5K5W3ZYsSRG4F7/exec',
  
  // 當前環境 ('test' 或 'production')
  ENVIRONMENT: 'test',
  
  // 檔案大小限制 (bytes)
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  
  // 支援的檔案格式
  ALLOWED_FILE_TYPES: ['.xlsx', '.xls'],
  
  // 目標工作表名稱 (可在此設定預設值，或由使用者選擇)
  TARGET_SHEET_NAME: '11501',
  
  // Google Sheets 目標
  TARGET_GOOGLE_SHEET_NAME: 'SalarySystemTest',
  TARGET_GOOGLE_SHEET_TAB: '國安班表'
};

// 取得當前環境的 GAS URL
CONFIG.GAS_URL = CONFIG.ENVIRONMENT === 'test' ? CONFIG.GAS_URL_TEST : CONFIG.GAS_URL_PROD;
