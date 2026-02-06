/**
 * Jest Setup - 模擬 Google Apps Script 環境
 */

// 模擬 SpreadsheetApp
global.SpreadsheetApp = {
  openById: jest.fn(),
  getActiveSpreadsheet: jest.fn(),
  create: jest.fn()
};

// 模擬 ContentService
global.ContentService = {
  createTextOutput: jest.fn((content) => ({
    setMimeType: jest.fn().mockReturnThis(),
    setContent: jest.fn().mockReturnThis(),
    getContent: jest.fn(() => content)
  })),
  MimeType: {
    JSON: 'application/json',
    TEXT: 'text/plain',
    JAVASCRIPT: 'text/javascript'
  }
};

// 模擬 PropertiesService
const mockProperties = {};
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn((key) => mockProperties[key]),
    setProperty: jest.fn((key, value) => {
      mockProperties[key] = value;
    }),
    getProperties: jest.fn(() => ({ ...mockProperties })),
    deleteProperty: jest.fn((key) => {
      delete mockProperties[key];
    }),
    deleteAllProperties: jest.fn(() => {
      Object.keys(mockProperties).forEach(key => delete mockProperties[key]);
    })
  }))
};

// 模擬 Utilities
global.Utilities = {
  newBlob: jest.fn(),
  base64Decode: jest.fn(),
  base64Encode: jest.fn(),
  formatDate: jest.fn((date, timezone, format) => {
    return date.toISOString();
  }),
  sleep: jest.fn(),
  parseCsv: jest.fn(),
  jsonParse: jest.fn((str) => JSON.parse(str)),
  jsonStringify: jest.fn((obj) => JSON.stringify(obj))
};

// 模擬 Logger
global.Logger = {
  log: jest.fn(console.log)
};

// 模擬 Session
global.Session = {
  getEffectiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  })),
  getActiveUser: jest.fn(() => ({
    getEmail: jest.fn(() => 'test@example.com')
  })),
  getTimeZone: jest.fn(() => 'Asia/Taipei')
};

// 模擬 DriveApp
global.DriveApp = {
  getFileById: jest.fn(),
  createFile: jest.fn()
};

// 模擬 UrlFetchApp
global.UrlFetchApp = {
  fetch: jest.fn()
};

// 模擬 LockService
global.LockService = {
  getScriptLock: jest.fn(() => ({
    tryLock: jest.fn(() => true),
    releaseLock: jest.fn(),
    hasLock: jest.fn(() => true),
    waitLock: jest.fn()
  }))
};

// 模擬 CacheService
global.CacheService = {
  getScriptCache: jest.fn(() => ({
    get: jest.fn(),
    put: jest.fn(),
    remove: jest.fn(),
    removeAll: jest.fn()
  }))
};

// 清理函數（每個測試後執行）
afterEach(() => {
  jest.clearAllMocks();
});
