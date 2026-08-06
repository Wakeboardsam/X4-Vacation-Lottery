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
    assert.ok(stateRes.data['Schema Version']);
});
