const { test, beforeEach } = require('node:test');
const assert = require('node:assert/strict');
require('./gas-mock-ext');

const { submitVacation, getRosterForVacation, getAdminOptions, getWeekAvailability } = require('../Vacation.gs');
const { writeConfigState, readConfigState } = require('../State.gs');
const { initializeOrUpdateWorkbook } = require('../Schema.gs');
const { beginVacationRound1, endVacationEarly } = require('../Admin.gs');

beforeEach(() => {
    global.resetMockData();
    initializeOrUpdateWorkbook();
});

test('Vacation: Roster uses override or default target', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', 'pid1', '1234', '', 'TRUE', '1', '1', 'TRUE', '', '']);
    tm.appendRow(['Test2', 'pid2', '5678', '', 'TRUE', '2', '2', 'TRUE', '5', '']);

    const roster = getRosterForVacation('2025', 9);
    assert.equal(roster[0].target, 9);
    assert.equal(roster[1].target, 5);
});

test('Vacation: Roster uses override or default capacity', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Week Availability');
    sheet.appendRow(['2025-01-01', 'Prime', 'None', '2', '', '', '', '']);
    sheet.appendRow(['2025-01-08', 'Non-Prime', 'None', '', '', '', '', '']);

    const { data, map } = getWeekAvailability();
    assert.equal(data.length, 3); // header + 2 rows
    assert.equal(data[1][map['Capacity Override']], '2');
    assert.equal(data[2][map['Capacity Override']], '');
});

test('Vacation: End to End E2E Simulation Merge Gate', () => {
    // Merge gate: deterministic end-to-end simulation that starts Round 1, completes seniority pass,
    // automatically enters Round 2, continues serpentine, applies skips, and completes vacation.

    // Seed State
    writeConfigState({
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Phase': 'SETUP',
        'Active Year': '2025'
    });

    // Seed Roster (3 participants, varying targets)
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    // P1: Target 2. P2: Target 3. P3: Target 2.
    tm.appendRow(['P1', '1111', '1111', '', 'TRUE', '1', '3', 'TRUE', '2', '']);
    tm.appendRow(['P2', '2222', '2222', '', 'TRUE', '2', '2', 'TRUE', '3', '']);
    tm.appendRow(['P3', '3333', '3333', '', 'TRUE', '3', '1', 'TRUE', '2', '']);

    // Seed Availability (5 weeks, cap 2)
    const wa = ss.getSheetByName('Week Availability');
    wa.appendRow(['W1', 'Non-Prime', 'None', '2', '', '', '', '']);
    wa.appendRow(['W2', 'Non-Prime', 'None', '2', '', '', '', '']);
    wa.appendRow(['W3', 'Non-Prime', 'None', '2', '', '', '', '']);
    wa.appendRow(['W4', 'Non-Prime', 'None', '2', '', '', '', '']);
    wa.appendRow(['W5', 'Non-Prime', 'None', '2', '', '', '', '']);

    // Auth Bypass
    global.resolveAdminSession_ = () => true;
    global.resolveParticipantSession_ = (token) => {
        if (token === 't1') return { participantId: '1111', name: 'P1' };
        if (token === 't2') return { participantId: '2222', name: 'P2' };
        if (token === 't3') return { participantId: '3333', name: 'P3' };
        return null;
    };

    try {
        // Action 1: Admin Starts
        const rStart = beginVacationRound1('admin');
        assert.equal(rStart.ok, true);

    // Check Phase
    let conf = readConfigState();
    assert.equal(conf['Current Phase'], 'VACATION_SENIORITY');
    assert.equal(conf['Current Vacation Round'], '1');

    // Active window should have all 3, but in Seniority Order: P1, P2, P3
    let aw = JSON.parse(conf['Current Active Window']);

    // Turn 1: P1 takes 2 Non-Prime (Skip triggered)
    let p1Turn = aw.find(t => t.participantId === '1111');
    const r1 = submitVacation('1111', p1Turn.turnId, ['W1', 'W2']);
    assert.equal(r1.ok, true);
    // P1 target is 2. So P1 is now COMPLETE_FOR_PHASE.

    // Turn 2: P2 takes 2 Non-Prime (Skip triggered)
    conf = readConfigState();
    aw = JSON.parse(conf['Current Active Window']);
    let p2Turn = aw.find(t => t.participantId === '2222');
    const r2 = submitVacation('2222', p2Turn.turnId, ['W3', 'W4']);
    assert.equal(r2.ok, true);
    // P2 target is 3. P2 should skip once.

    // Turn 3: P3 takes 1 Non-Prime (No Skip)
    conf = readConfigState();
    aw = JSON.parse(conf['Current Active Window']);
    let p3Turn = aw.find(t => t.participantId === '3333');
    const r3 = submitVacation('3333', p3Turn.turnId, ['W1']);
    assert.equal(r3.ok, true);

    // End of Round 1!
    // Automatic transition to Round 2 (VACATION_RANDOM)
    conf = readConfigState();
    assert.equal(conf['Current Phase'], 'VACATION_RANDOM');
    assert.equal(conf['Current Vacation Round'], '2');

    // P1 is complete.
    // P2 is skipping.
    // P3 is eligible.
    // Order in Lottery (Lottery POS: P3=1, P2=2, P1=3).
    // Serpentine Forward.
    // P3 is active, P2 skips, P1 excluded.
    aw = JSON.parse(conf['Current Active Window']);

    // The Queue Engine handles skips dynamically but depending on window rules it might advance through P2's skip and pick them up on the SAME turn or simply drop them.
    // In our manual QueueEngine, it should consume P2's skip and put P2 back into the active window if the window isn't full, OR drop P2.
    // Let's actually find P3's turn and advance.
    let p3Turn2 = aw.find(t => t.participantId === '3333');
    assert.ok(p3Turn2);

    // Turn 4: P3 takes 1 Non-Prime. Target is 2, so P3 becomes COMPLETE_FOR_PHASE
    const r4 = submitVacation('3333', p3Turn2.turnId, ['W2']);
    assert.equal(r4.ok, true);

    // Now P1 complete, P3 complete. P2 is only one left.
    // P2 target is 3, has 2 selections.
    // Cycle advances.
    conf = readConfigState();
    aw = JSON.parse(conf['Current Active Window']);

    // P2 is now active. (If P3 completed the serpentine path, P2 gets it next)
    assert.equal(aw.length, 1);
    assert.equal(aw[0].participantId, '2222');

    // Turn 5: P2 takes 1 week. Target is 3, becomes COMPLETE_FOR_PHASE.
    const r5 = submitVacation('2222', aw[0].turnId, ['W5']);
    assert.equal(r5.ok, true, r5.message);

        // Phase should now auto-complete
        conf = readConfigState();
        assert.equal(conf['Current Phase'], 'VACATION_RANDOM'); // Never changes from the state it was in.
        assert.equal(conf['Phase Ready State'], 'READY_WEEKEND');
    } finally {
        delete global.resolveParticipantSession_;
    }
});

test('Vacation: Atomic Rollback Behavior - failed write resets matrix', () => {
    writeConfigState({
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Phase': 'VACATION_SENIORITY',
        'Active Year': '2025',
        'Current Active Window': JSON.stringify([{
             turnId: 't1', participantId: '1111', cycle: 0, direction: 'FORWARD', position: 1
        }])
    });

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', '1111', '1111', '', 'TRUE', '1', '1', 'TRUE', '9', '']);

    const wa = ss.getSheetByName('Week Availability');
    wa.appendRow(['W1', 'Non-Prime', 'None', '2', '', '', '', '']);
    wa.appendRow(['W2', 'Non-Prime', 'None', '2', '', '', '', '']);

    global.resolveParticipantSession_ = () => ({ participantId: '1111', name: 'Test1' });

    // Force `writeConfigState` to fail internally by mocking it temporarily to throw
    const originalWriteConfigState = typeof writeConfigState_ === 'function' ? writeConfigState_ : require('../State.gs').writeConfigState;

    try {
        global.writeConfigState_ = () => { throw new Error('Simulated write failure'); };

        let threwFatalError = false;
        try {
            submitVacation('1111', 't1', ['W1', 'W2']);
        } catch(e) {
            threwFatalError = true;
            assert.match(e.message, /Rollback applied/);
        }
        assert.ok(threwFatalError, "Expected fatal rollback error to be thrown");

        // Verify assignments rolled back
        const { data, map } = getWeekAvailability();
        assert.equal(data[1][map['Person1']], '');
        assert.equal(data[2][map['Person1']], '');

        // Verify config state hasn't moved
        const conf = readConfigState();
        assert.equal(conf['Current Active Window'].includes('t1'), true);

    } finally {
        global.writeConfigState_ = originalWriteConfigState;
        delete global.resolveParticipantSession_;
        delete global.writeConfigState_; // Clears the manual override since tests use isolated dependencies
    }
});

test('Vacation: Rejects selections if not eligible', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    // Active but Vacation Phase Disabled
    tm.appendRow(['Alice', '1111', '1111', '', 'TRUE', '1', '1', 'FALSE', '', '']);

    writeConfigState({
        'Current Phase': 'VACATION_SENIORITY',
        'Current Active Window': JSON.stringify([{
             turnId: 't1', participantId: '1111', cycle: 0, direction: 'FORWARD', position: 1
        }])
    });

    global.resolveParticipantSession_ = () => ({ participantId: '1111', name: 'Alice' });
    try {
        const res = submitVacation('1111', 't1', ['2025-01-01']);
        assert.equal(res.ok, false);
        assert.match(res.message, /not eligible/);
    } finally {
        delete global.resolveParticipantSession_;
    }
});

test('Vacation: Successfully makes two Non-Prime selections and applies skip', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Bob', '2222', '2222', '', 'TRUE', '1', '1', 'TRUE', '9', '']);
    tm.appendRow(['Alice', '3333', '3333', '', 'TRUE', '2', '2', 'TRUE', '9', '']); // Need a second participant so round doesn't immediately complete and reset!

    const wa = ss.getSheetByName('Week Availability');
    wa.appendRow(['2025-01-01', 'Non-Prime', 'None', '4', '', '', '', '']);
    wa.appendRow(['2025-01-08', 'Non-Prime', 'None', '4', '', '', '', '']);

    writeConfigState({
        'Current Phase': 'VACATION_SENIORITY',
        'Current Queue Phase': 'VACATION_SENIORITY',
        'Current Vacation Round': '1',
        'Active Year': '2025',
        'Current Queue Cursor': '0',
        'Current Active Window': JSON.stringify([{
             turnId: 't2', participantId: '2222', cycle: 0, direction: 'FORWARD', position: 1
        }]),
        'Current Queue Skip State': '{}'
    });

    global.resolveParticipantSession_ = () => ({ participantId: '2222', name: 'Bob' });
    try {
        const res = submitVacation('2222', 't2', ['2025-01-01', '2025-01-08']);
        assert.equal(res.ok, true, res.message);

        const config = readConfigState();
        const skipState = JSON.parse(config['Current Queue Skip State'] || '{}');
        assert.equal(skipState['2222'], 1);
    } finally {
        delete global.resolveParticipantSession_;
    }
});

test('Vacation: Skips are preserved across Round 1 -> Round 2 transition', () => {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Bob', '2222', '2222', '', 'TRUE', '1', '1', 'TRUE', '9', '']);
    // No Alice here, so Bob finishing will trigger the Round 1 -> 2 transition

    const wa = ss.getSheetByName('Week Availability');
    wa.appendRow(['2025-01-01', 'Non-Prime', 'None', '4', '', '', '', '']);
    wa.appendRow(['2025-01-08', 'Non-Prime', 'None', '4', '', '', '', '']);

    writeConfigState({
        'Current Phase': 'VACATION_SENIORITY',
        'Current Queue Phase': 'VACATION_SENIORITY',
        'Current Vacation Round': '1',
        'Active Year': '2025',
        'Current Queue Cursor': '0',
        'Current Active Window': JSON.stringify([{
             turnId: 't2', participantId: '2222', cycle: 0, direction: 'FORWARD', position: 1
        }]),
        'Current Queue Skip State': '{}'
    });

    global.resolveParticipantSession_ = () => ({ participantId: '2222', name: 'Bob' });
    try {
        const res = submitVacation('2222', 't2', ['2025-01-01', '2025-01-08']);
        assert.equal(res.ok, true, res.message);
    } finally {
        delete global.resolveParticipantSession_;
    }

    const config = readConfigState();
    const skipState = JSON.parse(config['Current Queue Skip State'] || '{}');
    assert.equal(skipState['2222'], 1);
    assert.equal(config['Current Phase'], 'VACATION_RANDOM');
    assert.equal(config['Current Vacation Round'], '2');
});

test('Admin: beginVacationRound1 success', () => {
    writeConfigState({
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Phase': 'SETUP',
        'Active Year': '2025'
    });
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tm = ss.getSheetByName('Turn Management');
    tm.appendRow(['Test1', '1234', '1234', '', 'TRUE', '1', '1', 'TRUE', '', '']);

    // Bypass auth for test
    global.resolveAdminSession_ = () => true;

    const res = beginVacationRound1('token');
    assert.equal(res.ok, true, res.message);
    const config = readConfigState();
    assert.equal(config['Current Phase'], 'VACATION_SENIORITY');
    assert.equal(config['Current Vacation Round'], '1');
});

test('Admin: endVacationEarly success', () => {
    writeConfigState({
        'Phase Ready State': 'READY_VACATION_SENIORITY',
        'Current Phase': 'VACATION_SENIORITY',
        'Active Year': '2025'
    });

    // Bypass auth for test
    global.resolveAdminSession_ = () => true;

    const res = endVacationEarly('token');
    assert.equal(res.ok, true, res.message);
    const config = readConfigState();
    assert.equal(config['Current Phase'], 'VACATION_SENIORITY'); // Doesn't change
    assert.equal(config['Phase Ready State'], 'READY_WEEKEND');
});
