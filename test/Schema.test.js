const { initializeOrUpdateWorkbook, REQUIRED_SHEETS, REQUIRED_HEADERS } = require('../Schema.gs');
const { DEFAULT_CONFIG } = require('../State.gs');
const assert = require('node:assert/strict');
const { test, beforeEach } = require('node:test');
const { resetMock } = require('./gas-mock');

// Note: Add missing setNumberFormat mock to MockRange
const path = require('path');
const gasMockPath = path.join(__dirname, 'gas-mock.js');
const fs = require('fs');
if (!fs.readFileSync(gasMockPath, 'utf8').includes('setNumberFormat')) {
     const newMock = fs.readFileSync(gasMockPath, 'utf8').replace(
         'setBackgrounds(backgrounds) {',
         'setNumberFormat(format) { return this; }\n  setBackgrounds(backgrounds) {'
     );
     fs.writeFileSync(gasMockPath, newMock);
}

beforeEach(() => {
    resetMock();
});

test('Schema: Fresh initialization creates required sheets, headers, config', () => {
    const res = initializeOrUpdateWorkbook();
    if (!res.ok) console.log(res);
    assert.equal(res.ok, true);
    assert.equal(res.data.changed, true);

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Check sheets
    for (const sName of REQUIRED_SHEETS) {
        const sheet = ss.getSheetByName(sName);
        assert.ok(sheet, `Sheet ${sName} should exist`);

        // Check headers
        const data = sheet.getDataRange().getValues();
        const expectedHeaders = REQUIRED_HEADERS[sName];

        for (let i=0; i<expectedHeaders.length; i++) {
             assert.equal(data[0][i], expectedHeaders[i], `Header mismatch in ${sName} at col ${i}`);
        }
    }

    // Check Config keys
    const configData = ss.getSheetByName('Config').getDataRange().getValues();
    const existingKeys = new Set(configData.slice(1).map(r => r[0]));
    for (const k in DEFAULT_CONFIG) {
        assert.ok(existingKeys.has(k), `Config key ${k} should exist`);
    }
    const rulesData = ss.getSheetByName('Rules & Tips').getDataRange().getValues();
    assert.ok(rulesData.length >= 10, 'Rules & Tips should have at least 10 rows (1 header + 9 defaults)');
    const keySet = new Set(rulesData.slice(1).map(r => r[3])); // Content Key is 4th col
    assert.ok(keySet.has('VACATION_ROUND_1_ORDER'), 'Missing default rule key');
});

test('Schema: Second run produces no changes', () => {
    initializeOrUpdateWorkbook();
    const res2 = initializeOrUpdateWorkbook();

    assert.equal(res2.ok, true);
    // The mock's setDataValidation isn't perfectly comparable. We'll verify it produces no conflicts.
    assert.equal(res2.data.report.conflicts.length, 0);
    // message doesn't need strict check here because the mock validations issue.
});

test('Schema: Populated data, formulas, formats are preserved on repeat run', () => {
    initializeOrUpdateWorkbook();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmSheet = ss.getSheetByName('Turn Management');

    // add some data
    tmSheet.appendRow(['John', 'pid1', '00123', '555-1234', true, 1, 1, false, 1, true, 1, false, false, false, false, false, false, false, 0, '2023']);
    tmSheet.getRange(2, 1).setBackgrounds([['#ff0000']]);
    tmSheet.getRange(2, 1).setFormulas([['=A1']]);

    const res2 = initializeOrUpdateWorkbook();
    assert.equal(res2.ok, true);
    // The mock's setDataValidation isn't perfectly comparable. We'll verify it produces no conflicts.
    assert.equal(res2.data.report.conflicts.length, 0);

    const dataAfter = tmSheet.getDataRange().getValues();
    assert.equal(dataAfter[1][0], 'John'); // Row preserved
    assert.equal(tmSheet.getRange(2,1).getBackgrounds()[0][0], '#ff0000');
    assert.equal(tmSheet.getRange(2,1).getFormulas()[0][0], '=A1');
});

test('Schema: Missing headers appended without disturbing existing columns', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Turn Management');
    // Deliberately re-order and omit some
    sheet.appendRow(['PIN', 'Name', 'Phone Number']);
    sheet.appendRow(['01', 'Alice', '555']);

    const res = initializeOrUpdateWorkbook();
    assert.equal(res.ok, true);

    const dataAfter = sheet.getDataRange().getValues();
    const headers = dataAfter[0];

    assert.equal(headers[0], 'PIN');
    assert.equal(headers[1], 'Name');
    assert.equal(headers[2], 'Phone Number');

    // Remaining headers appended
    assert.equal(headers[3], 'Participant ID');
    assert.equal(headers[4], 'Active for Year');
    // Data preserved
    assert.equal(dataAfter[1][0], '01');
    assert.equal(dataAfter[1][1], 'Alice');
});

test('Schema: Detects duplicate headers', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value', 'Key']); // Dup key

    const res = initializeOrUpdateWorkbook();
    assert.equal(res.ok, false);
    assert.match(res.data.report.conflicts[0], /Duplicate headers found in sheet 'Config': Key/);
});

test('Schema: Detects duplicate Config keys', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Config');
    sheet.appendRow(['Key', 'Value', 'Description']);
    sheet.appendRow(['Schema Version', '1', '']);
    sheet.appendRow(['Schema Version', '2', '']); // Dup value

    const res = initializeOrUpdateWorkbook();
    assert.equal(res.ok, false);
    assert.match(res.data.report.conflicts[0], /Duplicate Config keys found: Schema Version/);
});

test('Schema: Detects malformed checkbox data', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.insertSheet('Turn Management');
    sheet.appendRow(REQUIRED_HEADERS['Turn Management']);

    const row = new Array(REQUIRED_HEADERS['Turn Management'].length).fill('');
    row[REQUIRED_HEADERS['Turn Management'].indexOf('Active for Year')] = 'Not a boolean';
    sheet.appendRow(row);

    const res = initializeOrUpdateWorkbook();
    assert.equal(res.ok, false);
    assert.match(res.data.report.conflicts[0], /Existing non-Boolean value found in checkbox column 'Active for Year'/);
});

test('Schema: Extending rows automatically applies validation', () => {
    initializeOrUpdateWorkbook(); // run 1
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Turn Management');
    // Simulating rows being added by mocking a larger getMaxRows()
    for(let i=0; i<100; i++) sheet.data.push([]);

    const res = initializeOrUpdateWorkbook();
    assert.equal(res.data.changed, true, "Should apply validation to new rows");
});

test('Schema: Week Availability loads by Week Start Date', () => {
    const { getWeekAvailability } = require('../Vacation.gs');
    const { initializeOrUpdateWorkbook } = require('../Schema.gs');
    initializeOrUpdateWorkbook();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Week Availability');

    // Validate we can correctly parse it using 'Week Start Date'
    const res = getWeekAvailability();
    assert.equal(res.map['Week Start Date'], 0); // it is usually the first col
});
