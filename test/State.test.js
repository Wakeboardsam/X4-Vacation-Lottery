const { readConfigState, writeConfigState, validateConfigState, DEFAULT_CONFIG } = require('../State.gs');
const assert = require('node:assert/strict');
const { test, beforeEach } = require('node:test');
const { resetMock } = require('./gas-mock');

beforeEach(() => {
    resetMock();
});

test('State: readConfigState throws if sheet missing', () => {
    assert.throws(() => readConfigState(), /Config sheet is missing/);
});

test('State: readConfigState reads properly formatted sheet', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value', 'Description']);
    sheet.appendRow(['Schema Version', '1', '']);
    sheet.appendRow(['Current Phase', 'SETUP', '']);

    const config = readConfigState();
    assert.equal(config['Schema Version'], '1');
    assert.equal(config['Current Phase'], 'SETUP');
});

test('State: validateConfigState valid case', () => {
    assert.doesNotThrow(() => validateConfigState(DEFAULT_CONFIG));
});

test('State: validateConfigState invalid phase', () => {
    const badConfig = { ...DEFAULT_CONFIG, 'Current Phase': 'INVALID_PHASE' };
    assert.throws(() => validateConfigState(badConfig), /Invalid Current Phase/);
});

test('State: validateConfigState invalid queue-list', () => {
     const badConfig = { ...DEFAULT_CONFIG, 'Current Active Window': 'not-an-array' };
     assert.throws(() => validateConfigState(badConfig), /Invalid queue-list serialization/);
});

test('State: writeConfigState success', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value']);
    for (const k in DEFAULT_CONFIG) {
         sheet.appendRow([k, DEFAULT_CONFIG[k]]);
    }

    writeConfigState({
        'Current Phase': 'VACATION_SENIORITY',
        'Current Vacation Round': '1'
    });

    const updated = readConfigState();
    assert.equal(updated['Current Phase'], 'VACATION_SENIORITY');
    assert.equal(updated['Current Vacation Round'], '1');
    assert.equal(updated['Setup State'], 'SETUP_EMPTY'); // unchanged
});

test('State: writeConfigState fails on invalid update', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value']);
    for (const k in DEFAULT_CONFIG) {
         sheet.appendRow([k, DEFAULT_CONFIG[k]]);
    }

    assert.throws(() => {
        writeConfigState({ 'Current Phase': 'BOGUS' });
    }, /Invalid Current Phase: BOGUS/);

    // verify state unchanged
    const unchanged = readConfigState();
    assert.equal(unchanged['Current Phase'], 'SETUP');
});

test('State: simulate atomic write failure does not leave partial state', () => {
    // In our simplified mock, an error thrown mid-loop during setValue would leave partial state.
    // However, our code collects all valid updates *first* and then applies them,
    // ensuring validation catches issues before any writes.
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value']);
    for (const k in DEFAULT_CONFIG) {
         sheet.appendRow([k, DEFAULT_CONFIG[k]]);
    }

    // To properly simulate, we need a custom mock scenario if we wanted an internal error,
    // but the required test is that invalid data doesn't partially update.
    assert.throws(() => {
        writeConfigState({
            'Current Phase': 'VACATION_SENIORITY', // valid
            'Current Vacation Round': 'BOGUS'      // invalid
        });
    }, /Invalid numeric value/);

    // verify phase wasn't updated
    const unchanged = readConfigState();
    assert.equal(unchanged['Current Phase'], 'SETUP');
});
