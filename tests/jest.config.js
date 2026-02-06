module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>/../gas-backend'],
  testMatch: [
    '**/__tests__/**/*.test.js',
    '**/?(*.)+(spec|test).js'
  ],
  collectCoverageFrom: [
    '../gas-backend/**/*.gs',
    '../gas-backend/**/*.js',
    '!../gas-backend/**/__tests__/**',
    '!../gas-backend/**/node_modules/**'
  ],
  setupFilesAfterEnv: ['<rootDir>/jest.setup.js'],
  testPathIgnorePatterns: [
    '/node_modules/',
    '/.clasp.json',
    '/appsscript.json'
  ],
  coverageDirectory: '<rootDir>/coverage',
  verbose: true
};
