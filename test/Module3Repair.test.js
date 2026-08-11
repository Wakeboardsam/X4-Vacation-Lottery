const assert = require('node:assert/strict');
const { test, beforeEach, afterEach } = require('node:test');
require('./gas-mock-ext');

const Utils = require('../Utils.gs');
const Auth = require('../Auth.gs');
const Vacation = require('../Vacation.gs');
const Admin = require('../Admin.gs');
const Code = require('../Code.gs');
const { calculateNextQueueState } = require('../QueueEngine.gs');
const { initializeOrUpdateWorkbook } = require('../Schema.gs');
const { readConfigState, writeConfigState } = require('../State.gs');

function appendParticipant(values) {
    SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Turn Management').appendRow(values);
}

function seedVacationState(activeWindow) {
    writeConfigState({
        'Active Year': '2027',
        'Current Phase': 'VACATION_SENIORITY',
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Vacation Round': '1',
        'Current Queue Phase': 'VACATION_SENIORITY',
        'Current Queue Order Source': 'seniority',
        'Current Queue Cursor': '1',
        'Current Queue Cycle': '1',
        'Current Serpentine Direction': 'FORWARD',
        'Current Active Window': JSON.stringify(activeWindow),
        'Current Directional Window': '[]',
        'Current Directional Window Completed': '[]',
        'Current Queue Skip State': '{}'
    });
}

function installCodeGlobals() {
    global.apiResponse_ = Utils.apiResponse;
    global.resolveParticipantSession_ = Auth.resolveParticipantSession;
    global.getWeekAvailability_ = Vacation.getWeekAvailability;
    global.getAdminOptions_ = Vacation.getAdminOptions;
    global.getRosterForVacation_ = Vacation.getRosterForVacation;
    global.normalizeWeekStartDate_ = Vacation.normalizeWeekStartDate;
    global.readConfigState_ = readConfigState;
    global.submitVacation_ = Vacation.submitVacation;
}

function clearInjectedGlobals() {
    for (const key of [
        'apiResponse_', 'resolveParticipantSession_', 'getWeekAvailability_', 'getAdminOptions_',
        'getRosterForVacation_', 'normalizeWeekStartDate_', 'readConfigState_', 'submitVacation_',
        'writeConfigState_', 'resolveAdminSession_'
    ]) delete global[key];
}

beforeEach(() => {
    global.resetMockData();
    const init = initializeOrUpdateWorkbook();
    assert.equal(init.ok, true, init.message);
});

afterEach(clearInjectedGlobals);

test('Module 3 identity: session is keyed by stable Participant ID, not PIN', () => {
    appendParticipant(['Alice', 'employee-alpha', '0042', '', true, 1, 3, true, '', '']);

    const login = Auth.loginParticipant('0042');
    assert.equal(login.ok, true);
    assert.equal(login.data.participant.participantId, 'employee-alpha');

    const projection = Auth.resolveParticipantSession(login.data.token);
    assert.equal(projection.participantId, 'employee-alpha');
    assert.notEqual(projection.participantId, '0042');

    const stored = Object.values(PropertiesService.getScriptProperties().getProperties()).join(' ');
    assert.match(stored, /employee-alpha/);
    assert.equal(stored.includes('"pin":"0042"'), false);
});

test('Module 3 identity: blank and duplicate Participant IDs fail closed without PIN fallback', () => {
    appendParticipant(['Blank ID', '', '1111', '', true, 1, 1, true, '', '']);
    let result = Auth.loginParticipant('1111');
    assert.equal(result.ok, false);
    assert.match(result.message, /participant roster requires administrator correction/);

    global.resetMockData();
    initializeOrUpdateWorkbook();
    appendParticipant(['Alpha', 'stable-id', '1111', '', true, 1, 1, true, '', '']);
    appendParticipant(['Beta', 'stable-id', '2222', '', true, 2, 2, true, '', '']);
    result = Auth.loginParticipant('1111');
    assert.equal(result.ok, false);
    assert.match(result.message, /participant roster requires administrator correction/);
    assert.throws(() => Vacation.getRosterForVacation('2027', 9), /duplicate Participant ID/);
});

test('Module 3 API: authenticated snapshot and submission use server session Participant ID', () => {
    appendParticipant(['Alice', 'employee-alpha', '0042', '', true, 1, 1, true, 1, '']);
    const wa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Week Availability');
    wa.appendRow([new Date(2027, 0, 4), 'Non-Prime', 'None', '', '', '', '', '']);
    seedVacationState([{ turnId: 'turn-alpha', participantId: 'employee-alpha', cycle: 1, direction: 'FORWARD', position: 0 }]);

    const login = Auth.loginParticipant('0042');
    assert.equal(login.ok, true);
    installCodeGlobals();

    const snapshot = Code.apiGetParticipantVacationData(login.data.token);
    assert.equal(snapshot.ok, true, snapshot.message);
    assert.equal(snapshot.data.turn.participantId, 'employee-alpha');
    assert.equal(snapshot.data.weeks[0].weekId, '2027-01-04');

    const submitted = Code.apiSubmitVacation(login.data.token, 'turn-alpha', ['2027-01-04']);
    assert.equal(submitted.ok, true, submitted.message);
    const data = Vacation.getWeekAvailability();
    assert.equal(data.data[1][data.map.Person1], 'Alice');
});

test('Module 3 ordering: shuffled rows use Seniority in Round 1 and Lottery in Round 2+', () => {
    const roster = [
        { participantId: 'C', seniority: 3, lottery: 1, disposition: 'ELIGIBLE' },
        { participantId: 'A', seniority: 1, lottery: 2, disposition: 'ELIGIBLE' },
        { participantId: 'B', seniority: 2, lottery: 3, disposition: 'ELIGIBLE' }
    ];
    const initial = {
        'Active Year': '2027', 'Current Queue Skip State': {}, 'Current Active Window': [],
        'Current Directional Window': [], 'Current Directional Window Completed': [], 'Active Window Generation': 0
    };

    const seniority = calculateNextQueueState(initial, roster, {
        action: 'INIT', windowSize: 3, phase: 'VACATION_SENIORITY', orderSource: 'seniority', movementMode: 'FORWARD_ONLY'
    });
    assert.deepEqual(seniority.nextState['Current Active Window'].map(turn => turn.participantId), ['A', 'B', 'C']);

    const lottery = calculateNextQueueState(initial, roster, {
        action: 'INIT', windowSize: 3, phase: 'VACATION_RANDOM', orderSource: 'lottery', movementMode: 'SERPENTINE'
    });
    assert.deepEqual(lottery.nextState['Current Active Window'].map(turn => turn.participantId), ['C', 'A', 'B']);
    assert.deepEqual(roster.map(person => person.participantId), ['C', 'A', 'B'], 'pure engine must not mutate caller roster');
});

test('Module 3 ordering: administrator start path ignores Turn Management row order', () => {
    appendParticipant(['C', 'id-c', '3003', '', true, 3, 1, true, 9, '']);
    appendParticipant(['A', 'id-a', '1001', '', true, 1, 2, true, 9, '']);
    appendParticipant(['B', 'id-b', '2002', '', true, 2, 3, true, 9, '']);
    writeConfigState({
        'Active Year': '2027', 'Current Phase': 'SETUP', 'Phase Ready State': 'READY_VACATION_SENIORITY'
    });
    global.resolveAdminSession_ = () => true;

    const started = Admin.beginVacationRound1('admin');
    assert.equal(started.ok, true, started.message);
    const active = JSON.parse(readConfigState()['Current Active Window']);
    assert.deepEqual(active.map(turn => turn.participantId), ['id-a', 'id-b', 'id-c']);
});

test('Module 3 ordering: invalid and duplicate authoritative positions fail instead of using row order', () => {
    const initial = {
        'Active Year': '2027', 'Current Queue Skip State': {}, 'Current Active Window': [],
        'Current Directional Window': [], 'Current Directional Window Completed': [], 'Active Window Generation': 0
    };
    assert.throws(() => calculateNextQueueState(initial, [
        { participantId: 'A', seniority: 1, disposition: 'ELIGIBLE' },
        { participantId: 'B', seniority: 1, disposition: 'ELIGIBLE' }
    ], { action: 'INIT', windowSize: 1, phase: 'VACATION_SENIORITY', orderSource: 'seniority', movementMode: 'FORWARD_ONLY' }), /Duplicate seniority position/);

    assert.throws(() => calculateNextQueueState(initial, [
        { participantId: 'A', lottery: 0, disposition: 'ELIGIBLE' }
    ], { action: 'INIT', windowSize: 1, phase: 'VACATION_RANDOM', orderSource: 'lottery', movementMode: 'SERPENTINE' }), /positive whole numbers/);
});

test('Module 3 skip: pending skip excludes exactly one appearance, is consumed, then participant returns', () => {
    const roster = [
        { participantId: 'A', lottery: 1, disposition: 'ELIGIBLE' },
        { participantId: 'B', lottery: 2, disposition: 'ELIGIBLE' },
        { participantId: 'C', lottery: 3, disposition: 'ELIGIBLE' }
    ];
    const initial = {
        'Active Year': '2027', 'Current Queue Skip State': { B: 1 }, 'Current Active Window': [],
        'Current Directional Window': [], 'Current Directional Window Completed': [], 'Active Window Generation': 0
    };
    let result = calculateNextQueueState(initial, roster, {
        action: 'INIT', preserveSkips: true, windowSize: 3, phase: 'VACATION_RANDOM', orderSource: 'lottery', movementMode: 'SERPENTINE'
    });
    let state = result.nextState;
    assert.deepEqual(state['Current Active Window'].map(turn => turn.participantId), ['A', 'C']);
    assert.equal(state['Current Queue Skip State'].B, undefined);

    for (const participantId of ['A', 'C']) {
        const turn = state['Current Active Window'].find(item => item.participantId === participantId);
        result = calculateNextQueueState(state, roster, {
            action: 'COMPLETE', completedTurnId: turn.turnId, windowSize: 3, orderSource: 'lottery', movementMode: 'SERPENTINE'
        });
        state = result.nextState;
    }
    assert.ok(state['Current Active Window'].some(turn => turn.participantId === 'B'), 'B returns in the following directional appearance');
});

test('Module 3 skip: a concurrently ACTIVE participant does not consume a pending skip', () => {
    const state = {
        'Active Year': '2027', 'Current Queue Phase': 'VACATION_RANDOM', 'Current Queue Order Source': 'lottery',
        'Current Queue Cycle': 1, 'Current Serpentine Direction': 'FORWARD', 'Current Queue Cursor': 0,
        'Current Queue Skip State': { A: 1 },
        'Current Active Window': [{ turnId: 'active-a', participantId: 'A', cycle: 1, direction: 'FORWARD', position: 0 }],
        'Current Directional Window': [], 'Current Directional Window Completed': [], 'Active Window Generation': 1
    };
    const result = calculateNextQueueState(state, [
        { participantId: 'A', lottery: 1, disposition: 'ELIGIBLE' },
        { participantId: 'B', lottery: 2, disposition: 'ELIGIBLE' }
    ], { action: 'RECONCILE', windowSize: 2, orderSource: 'lottery', movementMode: 'SERPENTINE' });
    assert.equal(result.nextState['Current Queue Skip State'].A, 1);
    assert.deepEqual(result.nextState['Current Active Window'].map(turn => turn.participantId), ['A', 'B']);
});

test('Module 3 rollback: Config failure restores assignments and leaves queue unchanged', () => {
    appendParticipant(['Alice', 'employee-alpha', '0042', '', true, 1, 1, true, 9, '']);
    const wa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Week Availability');
    wa.appendRow(['2027-01-04', 'Non-Prime', 'None', 2, '', '', '', '']);
    wa.appendRow(['2027-01-11', 'Non-Prime', 'None', 2, '', '', '', '']);
    seedVacationState([{ turnId: 'turn-alpha', participantId: 'employee-alpha', cycle: 1, direction: 'FORWARD', position: 0 }]);
    const before = readConfigState();

    global.writeConfigState_ = () => { throw new Error('injected Config failure'); };
    const result = Vacation.submitVacation('employee-alpha', 'turn-alpha', ['2027-01-04', '2027-01-11']);
    assert.equal(result.ok, false);
    assert.equal(result.message, 'Your vacation selection could not be saved. No changes were made. Please try again.');
    const weeks = Vacation.getWeekAvailability();
    assert.equal(weeks.data[1][weeks.map.Person1], '');
    assert.equal(weeks.data[2][weeks.map.Person1], '');
    assert.deepEqual(readConfigState(), before);
});

test('Module 3 rollback: rollback failure is the only fatal consistency error', () => {
    appendParticipant(['Alice', 'employee-alpha', '0042', '', true, 1, 1, true, 9, '']);
    const wa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Week Availability');
    wa.appendRow(['2027-01-04', 'Non-Prime', 'None', 2, '', '', '', '']);
    wa.appendRow(['2027-01-11', 'Non-Prime', 'None', 2, '', '', '', '']);
    seedVacationState([{ turnId: 'turn-alpha', participantId: 'employee-alpha', cycle: 1, direction: 'FORWARD', position: 0 }]);

    global.writeConfigState_ = () => { throw new Error('injected Config failure'); };
    const originalGetRange = wa.getRange.bind(wa);
    let matrixWrites = 0;
    wa.getRange = function(row, col, numRows, numCols) {
        const range = originalGetRange(row, col, numRows, numCols);
        const originalSetValues = range.setValues.bind(range);
        range.setValues = values => {
            if (row === 1 && col === 1 && numRows === 3) {
                matrixWrites++;
                if (matrixWrites === 2) throw new Error('injected rollback failure');
            }
            return originalSetValues(values);
        };
        return range;
    };

    assert.throws(() => Vacation.submitVacation('employee-alpha', 'turn-alpha', ['2027-01-04', '2027-01-11']), /Fatal consistency error/);
    wa.getRange = originalGetRange;
});

test('Module 3 schema: Capacity Override remains blank-or-positive and window settings remain required-positive', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const weekSheet = ss.getSheetByName('Week Availability');
    const weekHeaders = Utils.getHeaderMap(weekSheet.getDataRange().getValues());
    const capacityRule = weekSheet.getRange(2, weekHeaders['Capacity Override'] + 1).getDataValidation();
    assert.match(capacityRule.args[0], /ISBLANK/);
    assert.match(capacityRule.args[0], />0/);

    const adminSheet = ss.getSheetByName('Admin Options');
    const adminData = adminSheet.getDataRange().getValues();
    const adminMap = Utils.getHeaderMap(adminData);
    const windowRow = adminData.findIndex(row => row[adminMap.Setting] === 'Vacation ACTIVE-window size');
    assert.equal(adminData[windowRow][adminMap.Value], '3');
    const windowRule = adminSheet.getRange(windowRow + 1, adminMap.Value + 1).getDataValidation();
    assert.doesNotMatch(windowRule.args[0], /ISBLANK/);
    assert.match(windowRule.args[0], />0/);

    const repeat = initializeOrUpdateWorkbook();
    assert.equal(repeat.ok, true, repeat.message);
    assert.equal(adminSheet.getDataRange().getValues()[windowRow][adminMap.Value], '3');

    weekSheet.appendRow(['2027-01-04', 'Non-Prime', 'None', 0, '', '', '', '']);
    const conflict = initializeOrUpdateWorkbook();
    assert.equal(conflict.ok, false);
    assert.match(conflict.data.report.conflicts.join(' '), /Capacity Override|positive integer/);
});

test('Module 3 window configuration: sizes 1, 2, 3, and 5 apply; invalid values fail without state mutation', () => {
    const roster = Array.from({ length: 5 }, (_, index) => ({
        participantId: `P${index + 1}`, seniority: index + 1, disposition: 'ELIGIBLE'
    }));
    for (const size of [1, 2, 3, 5]) {
        const state = {
            'Active Year': '2027', 'Current Queue Skip State': {}, 'Current Active Window': [],
            'Current Directional Window': [], 'Current Directional Window Completed': [], 'Active Window Generation': 0
        };
        const result = calculateNextQueueState(state, roster, {
            action: 'INIT', windowSize: size, phase: 'VACATION_SENIORITY', orderSource: 'seniority', movementMode: 'FORWARD_ONLY'
        });
        assert.equal(result.nextState['Current Active Window'].length, size);
    }

    const adminSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Admin Options');
    const data = adminSheet.getDataRange().getValues();
    const map = Utils.getHeaderMap(data);
    const row = data.findIndex(item => item[map.Setting] === 'Vacation ACTIVE-window size');
    adminSheet.getRange(row + 1, map.Value + 1).setValue(0);
    const before = readConfigState();
    assert.throws(() => Vacation.getAdminOptions(), /positive whole number/);
    assert.deepEqual(readConfigState(), before);
});

test('Module 3 completion: ending vacation marks Weekend ready, keeps vacation phase closed, and does not start Weekend', () => {
    writeConfigState({
        'Active Year': '2027', 'Current Phase': 'VACATION_RANDOM', 'Phase Ready State': 'READY_VACATION_RANDOM',
        'Current Active Window': '[]'
    });
    global.resolveAdminSession_ = () => true;
    const ended = Admin.endVacationEarly('admin');
    assert.equal(ended.ok, true, ended.message);
    const config = readConfigState();
    assert.equal(config['Current Phase'], 'VACATION_RANDOM');
    assert.equal(config['Phase Ready State'], 'READY_WEEKEND');
    assert.notEqual(config['Current Phase'], 'WEEKEND');

    const rejected = Vacation.submitVacation('any-id', 'stale-turn', ['2027-01-04']);
    assert.equal(rejected.ok, false);
    assert.match(rejected.message, /Vacation selection is closed/);
});

test('Module 3 completion: natural Round 1 completion with all targets met marks Weekend ready', () => {
    appendParticipant(['Alice', 'id-a', '1001', '', true, 1, 2, true, 1, '']);
    appendParticipant(['Bob', 'id-b', '2002', '', true, 2, 1, true, 1, '']);
    const wa = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Week Availability');
    wa.appendRow(['2027-01-04', 'Non-Prime', 'None', 2, '', '', '', '']);
    wa.appendRow(['2027-01-11', 'Non-Prime', 'None', 2, '', '', '', '']);
    writeConfigState({
        'Active Year': '2027', 'Current Phase': 'SETUP', 'Phase Ready State': 'READY_VACATION_SENIORITY'
    });
    global.resolveAdminSession_ = () => true;
    assert.equal(Admin.beginVacationRound1('admin').ok, true);

    let active = JSON.parse(readConfigState()['Current Active Window']);
    assert.equal(Vacation.submitVacation('id-a', active.find(turn => turn.participantId === 'id-a').turnId, ['2027-01-04']).ok, true);
    active = JSON.parse(readConfigState()['Current Active Window']);
    assert.equal(Vacation.submitVacation('id-b', active.find(turn => turn.participantId === 'id-b').turnId, ['2027-01-11']).ok, true);

    const config = readConfigState();
    assert.equal(config['Current Phase'], 'VACATION_RANDOM');
    assert.equal(config['Phase Ready State'], 'READY_WEEKEND');
    assert.deepEqual(JSON.parse(config['Current Active Window']), []);
});
