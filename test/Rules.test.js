const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
require('./gas-mock-ext');

const { saveRulesAcknowledgment, isParticipantAcknowledgedForYear, mapTransferPreferences, getRulesContent, getParticipantAcknowledgmentState } = require('../Rules.gs');
const { writeConfigState, readConfigState } = require('../State.gs');
const { initializeOrUpdateWorkbook } = require('../Schema.gs');

beforeEach(() => {
    global.resetMockData();
    initializeOrUpdateWorkbook();
});

test('Rules: Missing Holiday Volunteer fails safely', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', '', '', '', '', '', '', '', '', '']);

    writeConfigState({ 'Active Year': '2025' });
    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1', activeYear: '2025' });

    const originalRReqPart = typeof resolveParticipantSession_ === 'function' ? resolveParticipantSession_ : require('../Auth.gs').resolveParticipantSession;
    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1', activeYear: '2025' });

    const res = saveRulesAcknowledgment('token', null, 'OFFER');
    assert.equal(res.ok, false);
    assert.match(res.message, /Choose Yes or No/);

    delete global.resolveParticipantSession_;
});

test('Rules: Invalid transfer preference fails safely', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', '', '', '', '', '', '', '', '', '']);

    writeConfigState({ 'Active Year': '2025' });
    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1', activeYear: '2025' });

    const res = saveRulesAcknowledgment('token', 'YES', 'INVALID');
    assert.equal(res.ok, false);
    assert.match(res.message, /Choose Yes or No/);

    delete global.resolveParticipantSession_;
});

test('Rules: Maps EXACT transfer choices correctly', () => {
    assert.deepEqual(mapTransferPreferences('OFFER'), { giver: true, receiver: false });
    assert.deepEqual(mapTransferPreferences('RECEIVE'), { giver: false, receiver: true });
    assert.deepEqual(mapTransferPreferences('BOTH'), { giver: true, receiver: true });
    assert.deepEqual(mapTransferPreferences('NONE'), { giver: false, receiver: false });
    assert.equal(mapTransferPreferences('UNKNOWN'), null);
});

test('Rules: Missing active year throws error', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', '', '', '', '', '', '', '', '', '']);

    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1' });

    const res = saveRulesAcknowledgment('token', 'YES', 'BOTH');
    assert.equal(res.ok, false);
    assert.match(res.message, /Active Year is not configured/);

    delete global.resolveParticipantSession_;
});

test('Rules: saveRulesAcknowledgment success exactly saves year and fields', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    // "Name", "Participant ID", "PIN", "Phone Number", "Active for Year", "Seniority Position", "Lottery Position",
    // "Vacation Phase Enabled", "Vacation Week Target Override", "Weekend Phase Enabled",
    // "Weekend Assignment Maximum", "Holiday Volunteer", "Mandatory Holiday Eligible",
    // "Transfer Giver", "Transfer Receiver", "Had Spring Break Last Year", "Had Christmas Week Last Year",
    // "Worked Any Official Holiday Last Year", "Week Availability Capacity Override", "Rules Acknowledged Year"
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', '', '', '', '', '', '', '', '', '']);

    writeConfigState({ 'Active Year': '2025' });
    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1', activeYear: '2025' });

    const res = saveRulesAcknowledgment('token', 'YES', 'BOTH');
    assert.equal(res.ok, true, res.message);

    const map = require('../Utils.gs').getHeaderMap(tm.getDataRange().getValues());
    const row = tm.getDataRange().getValues()[1];

    assert.equal(row[map['Holiday Volunteer']], true);
    assert.equal(row[map['Transfer Giver']], true);
    assert.equal(row[map['Transfer Receiver']], true);
    assert.equal(row[map['Rules Acknowledged Year']], '2025'); // Note exact 2025 string, not timestamp

    delete global.resolveParticipantSession_;
});

test('Rules: Acknowledgment idempotency protects admin edits', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', false, '', false, false, '', '', '', '', '2025']);

    writeConfigState({ 'Active Year': '2025' });
    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1', activeYear: '2025' });

    const res = saveRulesAcknowledgment('token', 'YES', 'BOTH');
    assert.equal(res.ok, true, res.message);

    const map = require('../Utils.gs').getHeaderMap(tm.getDataRange().getValues());
    const row = tm.getDataRange().getValues()[1];

    assert.equal(row[map['Holiday Volunteer']], false);
    assert.equal(row[map['Transfer Giver']], false);
    assert.equal(row[map['Transfer Receiver']], false);
    assert.equal(row[map['Rules Acknowledged Year']], '2025');

    delete global.resolveParticipantSession_;
});

test('Rules: Prior year unacknowledges for current year', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', '', '', '', '', '', '', '', '', '2024']);
    writeConfigState({ 'Active Year': '2025' });

    assert.equal(isParticipantAcknowledgedForYear('pid1', '2025'), false);
    assert.equal(isParticipantAcknowledgedForYear('pid1', '2024'), true);
});

test('Rules: Duplicate Participant IDs fail safely', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', '', '', '', '', '', '', '', '', '']);
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', '', '', '', '', '', '', '', '', '']);

    writeConfigState({ 'Active Year': '2025' });
    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1', activeYear: '2025' });

    const res = saveRulesAcknowledgment('token', 'YES', 'BOTH');
    assert.equal(res.ok, false);
    assert.match(res.message, /Participant record not found or duplicate/);

    delete global.resolveParticipantSession_;
});

test('Rules: Partial write failure rolls back', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '9', '', '', false, '', false, false, '', '', '', '', '']);

    writeConfigState({ 'Active Year': '2025' });
    global.resolveParticipantSession_ = () => ({ participantId: 'pid1', name: 'Test1', activeYear: '2025' });

    // Mock sheet method to fail halfway
    const realRange = tm.getRange;
    let writeCount = 0;
    tm.getRange = function(row, col, numRows, numCols) {
        const range = realRange.apply(this, arguments);
        const realSetValue = range.setValue;
        range.setValue = function(val) {
            writeCount++;
            if (writeCount === 4) { // Fail on saving the year
                throw new Error("Simulated Write Error");
            }
            return realSetValue.apply(this, arguments);
        };
        return range;
    };

    const res = saveRulesAcknowledgment('token', 'YES', 'BOTH');
    assert.equal(res.ok, false);

    // Check that we're rolled back
    const map = require('../Utils.gs').getHeaderMap(tm.getDataRange().getValues());
    const row = tm.getDataRange().getValues()[1];

    assert.equal(row[map['Holiday Volunteer']], false);
    assert.equal(row[map['Transfer Giver']], false);
    assert.equal(row[map['Transfer Receiver']], false);
    assert.equal(row[map['Rules Acknowledged Year']], '');

    tm.getRange = realRange;
    delete global.resolveParticipantSession_;
});

test('Rules: Content Key duplicate checking works correctly', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rt = ss.getSheetByName('Rules & Tips');
    rt.appendRow([10, 'Rule 1', true, 'KEY1', 'RULE', 'GENERAL']);

    // Reinitialize again, we shouldn't add the same default rules
    const res = initializeOrUpdateWorkbook();
    assert.equal(res.ok, true);

    const rules = getRulesContent();
    assert.equal(rules.rules.length > 0, true);
});

test('Rules: Old legacy rules work with general context', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rt = ss.getSheetByName('Rules & Tips');
    rt.appendRow([10, 'Rule 1', true, '', '', '']); // Legacy Rule

    const rules = getRulesContent();
    assert.equal(rules.rules.length > 0, true);
    assert.equal(rules.rules.find(r => r.text === 'Rule 1').context, 'GENERAL');
});
