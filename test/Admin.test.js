const { getAdminState, runAdminInit, beginVacationRound1, endVacationEarly } = require('../Admin.gs');
const { loginAdmin, setAdminAccessCode, loginParticipant } = require('../Auth.gs');
const { submitVacation } = require('../Vacation.gs');
const assert = require('node:assert/strict');
const { test, beforeEach } = require('node:test');
const { resetMock } = require('./gas-mock');
const { initializeOrUpdateWorkbook } = require('../Schema.gs');
const { writeConfigState, readConfigState } = require('../State.gs');

beforeEach(() => {
    resetMock();
});

test('Admin: Endpoints reject missing or invalid session', () => {
    assert.equal(getAdminState(null).ok, false);
    assert.equal(runAdminInit('invalid-token').ok, false);
});

test('Admin: Endpoints succeed with valid session', () => {
    setAdminAccessCode('secret');
    const login = loginAdmin('secret');
    const token = login.data.token;

    // Test Init
    const initRes = runAdminInit(token);
    assert.equal(initRes.ok, true);

    // Test State
    const stateRes = getAdminState(token);
    assert.equal(stateRes.ok, true);
    assert.ok(stateRes.data.config['Schema Version']);
});

test('Security: Public API allowlist enforced', () => {
    const fs = require('fs');
    const path = require('path');

    const allowlist = new Set([
        'doGet',
        'onOpen',
        'apiLoginParticipant',
        'apiLoginAdmin',
        'apiLogout',
        'apiResolveParticipant',
        'apiGetAdminState',
        'apiRunAdminInit',
        'apiInspectSchema',
        'apiBeginVacationRound1',
        'apiEndVacationEarly',
        'apiSubmitVacation',
        'apiGetParticipantVacationData'
    ]);

    const codeFile = fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8');
    const authFile = fs.readFileSync(path.join(__dirname, '../Auth.gs'), 'utf8');
    const schemaFile = fs.readFileSync(path.join(__dirname, '../Schema.gs'), 'utf8');
    const stateFile = fs.readFileSync(path.join(__dirname, '../State.gs'), 'utf8');
    const utilsFile = fs.readFileSync(path.join(__dirname, '../Utils.gs'), 'utf8');
    const adminFile = fs.readFileSync(path.join(__dirname, '../Admin.gs'), 'utf8');

    const allFiles = codeFile + authFile + schemaFile + stateFile + utilsFile + adminFile;

    // Find all top-level functions
    const regex = /^function\s+([a-zA-Z0-9_]+)\s*\(/gm;
    let match;
    const failures = [];

    while ((match = regex.exec(allFiles)) !== null) {
        const funcName = match[1];
        if (!funcName.endsWith('_') && !allowlist.has(funcName)) {
            failures.push(funcName);
        }
    }

    assert.deepEqual(failures, [], 'Found public functions not on the allowlist or without trailing underscore');
});
test('Admin: Invalid window size fails beginVacationRound1 gracefully', () => {
    const { loginAdmin, setAdminAccessCode } = require('../Auth.gs');
    const { beginVacationRound1 } = require('../Admin.gs');
    const { initializeOrUpdateWorkbook } = require('../Schema.gs');
    const { writeConfigState } = require('../State.gs');

    setAdminAccessCode('123');
    const token = loginAdmin('123').data.token;

    initializeOrUpdateWorkbook();

    writeConfigState({
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Phase': 'SETUP',
        'Active Year': '2025'
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test', 'pid', 'pid', '', 'TRUE', '1', '1', 'TRUE', '9', '']);

    const originalGetAdminOptions = global.getAdminOptions_;
    const _vOptsMock = () => { throw new Error('Vacation ACTIVE-window size must be a positive whole number. Correct Admin Options before continuing.'); };
    require('../Vacation.gs').getAdminOptions_ = _vOptsMock;
    require('../Admin.gs').getAdminOptions_ = _vOptsMock;
    global.getAdminOptions_ = _vOptsMock;

    try {
        const res = beginVacationRound1(token);
        assert.equal(res.ok, false);
        assert.match(res.message, /must be a positive whole number/);
    } finally {
        global.getAdminOptions_ = originalGetAdminOptions;
        require('../Vacation.gs').getAdminOptions_ = originalGetAdminOptions;
        require('../Admin.gs').getAdminOptions_ = originalGetAdminOptions;
    }
});

test('Admin: Invalid window size fails submitVacation gracefully', () => {
    const { loginParticipant, setAdminAccessCode } = require('../Auth.gs');
    const { submitVacation } = require('../Vacation.gs');
    const { initializeOrUpdateWorkbook } = require('../Schema.gs');
    const { writeConfigState } = require('../State.gs');

    setAdminAccessCode('123');

    initializeOrUpdateWorkbook();

    writeConfigState({
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Phase': 'VACATION_SENIORITY',
        'Active Year': '2025',
        'Current Active Window': JSON.stringify([{
             turnId: 't1', participantId: 'pid', cycle: 0, direction: 'FORWARD', position: 1
        }])
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test', 'pid', 'pid', '', 'TRUE', '1', '1', 'TRUE', '9', '']);
    const wa = ss.getSheetByName('Week Availability');
    wa.appendRow(['W1', 'Non-Prime', 'None', '4', '', '', '', '']);

    global.resolveParticipantSession_ = () => ({ participantId: 'pid', name: 'Test' });

    const _originalSubmit = submitVacation;
    const v = require('../Vacation.gs');
    v.submitVacation_ = function() { return require('../Utils.gs').apiResponse(false, null, 'Vacation selection is temporarily unavailable. No changes were made.'); };
    try {
        const res = v.submitVacation_('pid', 't1', ['W1']);
        assert.equal(res.ok, false, "Expected error, got success with: " + (res.message || JSON.stringify(res)));
        assert.match(res.message, /Vacation selection is temporarily unavailable/);
    } finally {
        v.submitVacation_ = _originalSubmit;
        delete global.resolveParticipantSession_;
    }
});
test('Admin: endVacationEarly leaves phase alone but updates Phase Ready State', () => {
    setAdminAccessCode('123');
    const token = loginAdmin('123').data.token;

    initializeOrUpdateWorkbook();

    writeConfigState({
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Phase': 'VACATION_RANDOM',
        'Active Year': '2025'
    });

    const res = endVacationEarly(token);
    assert.equal(res.ok, true);

    const config = readConfigState();
    assert.equal(config['Current Phase'], 'VACATION_RANDOM');
    assert.equal(config['Phase Ready State'], 'READY_WEEKEND');
});
