require('./gas-mock');
global.resetMockData = function() {
    global.SpreadsheetApp.reset();
    global.mockScriptProperties.deleteAllProperties();
    global.mockScriptLock.releaseLock();
};
