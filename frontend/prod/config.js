/**
 * 薪資計算系統 - 環境設定（正式環境）
 */

const CONFIG = {
  // GAS Web App URL
  GAS_URL_TEST: 'https://script.google.com/macros/s/AKfycbxpExUCOrTLcbUSHcaBmtluODqb2pNopEuh7QOUJXJXXVZcKMP_EJiUKRLsTTZ3sANS5Q/exec',
  GAS_URL_PROD: 'https://script.google.com/macros/s/185v2ySyq9ncyzBwbnnogQTXHuw3S78pfmXcaATxWRs5K5W3ZYsSRG4F7/exec',
  
  // 當前環境（正式）
  ENVIRONMENT: 'production',
  
  // 檔案大小限制 (bytes)
  MAX_FILE_SIZE: 10 * 1024 * 1024, // 10MB
  
  // 支援的檔案格式
  ALLOWED_FILE_TYPES: ['.xlsx', '.xls'],
  
  // 目標工作表名稱
  TARGET_SHEET_NAME: '11501',
  
  // Google Sheets 目標（正式環境由 GAS Script Properties 決定）
  TARGET_GOOGLE_SHEET_NAME: 'SalarySystem',
  TARGET_GOOGLE_SHEET_TAB: '國安班表'
};

CONFIG.GAS_URL = CONFIG.GAS_URL_PROD;
