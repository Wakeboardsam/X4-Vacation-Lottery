const { getAdminState, runAdminInit } = require('../Admin.gs');
const { loginAdmin, setAdminAccessCode } = require('../Auth.gs');
const assert = require('node:assert/strict');
const { test, beforeEach } = require('node:test');
const { resetMock } = require('./gas-mock');
const { initializeOrUpdateWorkbook } = require('../Schema.gs');

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
