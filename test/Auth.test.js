const {
    setAdminAccessCode, loginParticipant, resolveParticipantSession,
    loginAdmin, resolveAdminSession, logout, SESSION_PREFIX_USER, SESSION_PREFIX_ADMIN
} = require('../Auth.gs');
const { initializeOrUpdateWorkbook } = require('../Schema.gs');
const assert = require('node:assert/strict');
const { test, beforeEach } = require('node:test');
const { resetMock } = require('./gas-mock');

beforeEach(() => {
    resetMock();
    initializeOrUpdateWorkbook(); // Ensure schema exists for participant auth tests
});

test('Auth: Reject blank admin code in setup', () => {
    assert.throws(() => setAdminAccessCode(''), /Access code cannot be blank/);
});

test('Auth: Set admin code stores only salt and hash', () => {
    setAdminAccessCode('mysecret');
    const props = PropertiesService.getScriptProperties().getProperties();
    assert.ok(props['ADMIN_SALT']);
    assert.ok(props['ADMIN_HASH']);
    assert.equal(Object.values(props).includes('mysecret'), false);
});

test('Auth: Admin login success issues valid admin token and stores digest', () => {
    setAdminAccessCode('secret');
    const res = loginAdmin('secret');
    assert.equal(res.ok, true);
    assert.ok(res.data.token);

    // Verify it resolves
    assert.equal(resolveAdminSession(res.data.token), true);

    // Verify raw token is not stored
    const props = PropertiesService.getScriptProperties().getProperties();
    const tokenVals = Object.values(props);
    assert.equal(tokenVals.some(v => typeof v === 'string' && v.includes(res.data.token)), false);
});

test('Auth: Admin login fails with wrong code', () => {
    setAdminAccessCode('secret');
    const res = loginAdmin('wrong');
    assert.equal(res.ok, false);
});

test('Auth: Impersonation attempt via user token on admin endpoint fails', () => {
    setAdminAccessCode('secret');

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmSheet = ss.getSheetByName('Turn Management');
    tmSheet.appendRow(['Test', 'pid', '123', '', true, '', '']);

    const uRes = loginParticipant('123');
    const uToken = uRes.data.token;

    // Admin resolve should fail for user token
    assert.equal(resolveAdminSession(uToken), false);
});

test('Auth: Participant login exact match preserving leading zeroes', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmSheet = ss.getSheetByName('Turn Management');
    tmSheet.appendRow(['Zero', 'pid', '007', '', true, '', '']);

    // Test that number 7 does not match string 007
    const badRes = loginParticipant('7');
    assert.equal(badRes.ok, false);

    const goodRes = loginParticipant('007');
    assert.equal(goodRes.ok, true);
    assert.equal(goodRes.data.participant.name, 'Zero');
});

test('Auth: Duplicate PIN fails safe', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmSheet = ss.getSheetByName('Turn Management');
    tmSheet.appendRow(['A', 'pid1', '007', '', true, '', '']);
    tmSheet.appendRow(['B', 'pid2', '007', '', true, '', '']);

    const res = loginParticipant('007');
    assert.equal(res.ok, false);
    assert.match(res.message, /Invalid PIN/);
});

test('Auth: Resolve participant rereads fresh projection', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmSheet = ss.getSheetByName('Turn Management');
    const row = ['Bob', 'pid3', '555', '', true, '', '']; // active = true
    tmSheet.appendRow(row);

    const res = loginParticipant('555');
    const token = res.data.token;

    let proj = resolveParticipantSession(token);
    assert.equal(proj.isActive, true);

    // Mutate state in DB (Active is column 5 since we added Participant ID at column 2)
    tmSheet.getRange(tmSheet.getLastRow(), 5).setValue(false); // Make inactive

    proj = resolveParticipantSession(token);
    assert.equal(proj.isActive, false); // Picked up change
});

test('Auth: Safe response projection - PIN not returned', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmSheet = ss.getSheetByName('Turn Management');
    tmSheet.appendRow(['Safe', 'pid4', '999', '123-4567', true, '', '']);

    const res = loginParticipant('999');
    const projStr = JSON.stringify(res.data.participant);
    assert.equal(projStr.includes('999'), false, 'Should not leak PIN');
    assert.equal(projStr.includes('123-4567'), false, 'Should not leak phone');
});

test('Auth: Logout invalidates session', () => {
    setAdminAccessCode('123');
    const res = loginAdmin('123');
    const token = res.data.token;
    assert.equal(resolveAdminSession(token), true);

    logout(token, 'admin');
    assert.equal(resolveAdminSession(token), false);
});

test('Auth: Session resume fails closed on newly introduced duplicate Participant ID', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tmSheet = ss.getSheetByName('Turn Management');
    const row = ['DupeTest', 'pid5', '9999', '', true, '', ''];
    tmSheet.appendRow(row);

    const res = loginParticipant('9999');
    assert.equal(res.ok, true);

    // Add a duplicate stable identity with a different PIN.
    const dupeRow = ['Dupe2', 'pid5', '8888', '', true, '', ''];
    tmSheet.appendRow(dupeRow);

    assert.throws(() => resolveParticipantSession(res.data.token), /Participant roster identity conflict/);
});
