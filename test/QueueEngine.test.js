const assert = require('node:assert/strict');
const { test, beforeEach } = require('node:test');
const { calculateNextQueueState, processQueueMutation } = require('../QueueEngine.gs');
const { initializeOrUpdateWorkbook } = require('../Schema.gs');
const { writeConfigState } = require('../State.gs');
const { resetMock } = require('./gas-mock');

// We have our pure engine tests and the wrapper tests to run.

beforeEach(() => {
    resetMock();
});

test('calculateNextQueueState: Invalid window size', () => {
    assert.throws(() => calculateNextQueueState({}, [], { windowSize: 0 }), /Invalid window size/);
    assert.throws(() => calculateNextQueueState({}, [], { windowSize: -1 }), /Invalid window size/);
    assert.throws(() => calculateNextQueueState({}, [], { windowSize: 1.5 }), /Invalid window size/);
    assert.throws(() => calculateNextQueueState({}, [], { windowSize: 'two' }), /Invalid window size/);
});

test('calculateNextQueueState: Empty roster', () => {
    const initialState = {
        "Active Year": "2024",
        "Current Queue Skip State": {},
        "Current Active Window": [],
        "Current Directional Window": [],
        "Current Directional Window Completed": [],
        "Active Window Generation": 0
    };

    const { nextState, queueComplete, completionReason } = calculateNextQueueState(initialState, [], {
        action: 'INIT', windowSize: 2, phase: 'TEST_PHASE', orderSource: 'LOTTERY', movementMode: 'FORWARD_ONLY'
    });

    assert.equal(queueComplete, true);
    assert.equal(completionReason, 'EMPTY_ROSTER');
    assert.equal(nextState["Current Active Window"].length, 0);
});

test('calculateNextQueueState: Seniority Forward Pass completes', () => {
    const initialState = {
        "Active Year": "2024",
        "Current Queue Skip State": {},
        "Current Active Window": [],
        "Current Directional Window": [],
        "Current Directional Window Completed": [],
        "Active Window Generation": 0
    };
    const roster = [
        { participantId: 'p1', disposition: 'ELIGIBLE' },
        { participantId: 'p2', disposition: 'EXCLUDED' },
        { participantId: 'p3', disposition: 'ELIGIBLE' }
    ];

    // Init queue
    let result = calculateNextQueueState(initialState, roster, {
        action: 'INIT', windowSize: 1, phase: 'VACATION_SENIORITY', orderSource: 'SENIORITY', movementMode: 'FORWARD_ONLY'
    });

    assert.equal(result.queueComplete, false);
    assert.equal(result.nextState["Current Active Window"].length, 1);
    assert.equal(result.nextState["Current Active Window"][0].participantId, 'p1');
    assert.equal(result.nextState["Active Window Generation"], 1);

    // Complete p1 turn
    result = calculateNextQueueState(result.nextState, roster, {
        action: 'COMPLETE', completedTurnId: result.nextState["Current Active Window"][0].turnId,
        windowSize: 1, phase: 'VACATION_SENIORITY', orderSource: 'SENIORITY', movementMode: 'FORWARD_ONLY'
    });

    assert.equal(result.queueComplete, false);
    assert.equal(result.nextState["Current Active Window"].length, 1);
    assert.equal(result.nextState["Current Active Window"][0].participantId, 'p3'); // skips p2

    // Complete p3 turn
    result = calculateNextQueueState(result.nextState, roster, {
        action: 'COMPLETE', completedTurnId: result.nextState["Current Active Window"][0].turnId,
        windowSize: 1, phase: 'VACATION_SENIORITY', orderSource: 'SENIORITY', movementMode: 'FORWARD_ONLY'
    });

    assert.equal(result.queueComplete, true);
    assert.equal(result.completionReason, 'FORWARD_PASS_COMPLETE');
    assert.equal(result.nextState["Current Active Window"].length, 0);
});

test('calculateNextQueueState: Serpentine reversal strict barrier', () => {
    const initialState = {
        "Active Year": "2024",
        "Current Queue Phase": "VACATION_RANDOM",
        "Current Queue Cycle": 1,
        "Current Serpentine Direction": "FORWARD",
        "Current Queue Cursor": 0,
        "Current Queue Skip State": {},
        "Current Active Window": [],
        "Current Directional Window": [],
        "Current Directional Window Completed": [],
        "Active Window Generation": 0
    };
    const roster = [
        { participantId: 'p19', disposition: 'ELIGIBLE' },
        { participantId: 'p20', disposition: 'ELIGIBLE' },
        { participantId: 'p21', disposition: 'ELIGIBLE' }
    ];

    // READ action to fill initial window
    let result = calculateNextQueueState(initialState, roster, {
        action: 'RECONCILE', windowSize: 3, phase: 'VACATION_RANDOM', orderSource: 'LOTTERY', movementMode: 'SERPENTINE'
    });

    let st = result.nextState;
    assert.equal(st["Current Active Window"].length, 3);
    assert.equal(st["Current Directional Window"].length, 0);

    // Now window size is 3, cursor is 3 (end of roster).
    // Let's complete p21 first.
    let t21 = st["Current Active Window"].find(t => t.participantId === 'p21').turnId;
    result = calculateNextQueueState(st, roster, {
        action: 'COMPLETE', completedTurnId: t21, windowSize: 3, movementMode: 'SERPENTINE'
    });

    st = result.nextState;
    // Because cursor is at end, and movementMode is SERPENTINE, it should freeze the directional window.
    assert.equal(st["Current Directional Window"].length, 2); // 19 and 20 are frozen
    assert.equal(st["Current Directional Window Completed"].length, 0);
});

test('calculateNextQueueState: [19, 20, 21] reverse strict barrier full trace', () => {
    let state = {
        "Active Year": "2024",
        "Current Queue Phase": "VACATION_RANDOM",
        "Current Queue Cycle": 1,
        "Current Serpentine Direction": "FORWARD",
        "Current Queue Cursor": 0,
        "Current Queue Skip State": {},
        "Current Active Window": [],
        "Current Directional Window": [],
        "Current Directional Window Completed": [],
        "Active Window Generation": 1
    };
    // Let's pretend roster length is 3 and they are indexes 0, 1, 2. (representing 19, 20, 21)
    const roster = [
        { participantId: 'p19', disposition: 'ELIGIBLE' },
        { participantId: 'p20', disposition: 'ELIGIBLE' },
        { participantId: 'p21', disposition: 'ELIGIBLE' }
    ];

    // READ action to fill initial window
    let result = calculateNextQueueState(state, roster, {
        action: 'RECONCILE', windowSize: 3, movementMode: 'SERPENTINE'
    });

    state = result.nextState;
    assert.equal(state["Current Active Window"].length, 3);
    assert.equal(state["Current Directional Window"].length, 0); // Not frozen yet because we just loaded them.
    assert.equal(state["Current Queue Cursor"], 3);

    // Now, let's complete p21.
    const t21 = state["Current Active Window"].find(t => t.participantId === 'p21').turnId;
    result = calculateNextQueueState(state, roster, {
        action: 'COMPLETE', completedTurnId: t21, windowSize: 3, movementMode: 'SERPENTINE'
    });

    state = result.nextState;
    // Because it was end of line, the directional window is frozen.
    assert.equal(state["Current Directional Window"].length, 2); // 19 and 20 left
    assert.equal(state["Current Directional Window Completed"].length, 0);
    assert.equal(state["Current Active Window"].length, 2); // 19 and 20 left

    // Nobody from reverse enters yet!
    assert.equal(state["Current Serpentine Direction"], "FORWARD");

    // Complete p19
    const t19 = state["Current Active Window"].find(t => t.participantId === 'p19').turnId;
    result = calculateNextQueueState(state, roster, {
        action: 'COMPLETE', completedTurnId: t19, windowSize: 3, movementMode: 'SERPENTINE'
    });
    state = result.nextState;
    assert.equal(state["Current Directional Window Completed"].length, 1);
    assert.equal(state["Current Directional Window Completed"][0], t19);
    assert.equal(state["Current Serpentine Direction"], "FORWARD"); // Still waiting

    // Complete p20
    const t20 = state["Current Active Window"].find(t => t.participantId === 'p20').turnId;
    result = calculateNextQueueState(state, roster, {
        action: 'COMPLETE', completedTurnId: t20, windowSize: 3, movementMode: 'SERPENTINE'
    });
    state = result.nextState;

    // Reverse should have happened, new cycle started, new window opened!
    assert.equal(state["Current Serpentine Direction"], "BACKWARD");
    assert.equal(state["Current Queue Cycle"], 2);
    assert.equal(state["Current Active Window"].length, 3); // 21, 20, 19 in reverse order
    assert.equal(state["Current Active Window"][0].participantId, 'p21');
    assert.equal(state["Current Active Window"][1].participantId, 'p20');
    assert.equal(state["Current Active Window"][2].participantId, 'p19');
    assert.equal(state["Current Directional Window"].length, 0); // Barrier cleared
});


test('calculateNextQueueState: Idempotent READ does not advance generation', () => {
    let state = {
        "Active Year": "2024",
        "Current Queue Phase": "VACATION_RANDOM",
        "Current Queue Cycle": 1,
        "Current Serpentine Direction": "FORWARD",
        "Current Queue Cursor": 1,
        "Current Queue Skip State": {},
        "Current Active Window": [
            { turnId: '1', participantId: 'p1', cycle: 1, direction: 'FORWARD', position: 0 }
        ],
        "Current Directional Window": [],
        "Current Directional Window Completed": [],
        "Active Window Generation": 5
    };
    const roster = [
        { participantId: 'p1', disposition: 'ELIGIBLE' },
        { participantId: 'p2', disposition: 'ELIGIBLE' }
    ];

    // READ action should return exact state
    let result = calculateNextQueueState(state, roster, {
        action: 'READ', windowSize: 1, movementMode: 'SERPENTINE'
    });

    assert.equal(result.nextState["Active Window Generation"], 5);
    assert.equal(result.nextState, state); // exact object

    // RECONCILE with size 2 adds one person, generation should bump once
    result = calculateNextQueueState(state, roster, {
        action: 'RECONCILE', windowSize: 2, movementMode: 'SERPENTINE'
    });
    assert.equal(result.nextState["Active Window Generation"], 6);
    assert.equal(result.nextState["Current Active Window"].length, 2);
});

test('calculateNextQueueState: Duplicate or stale COMPLETE does not change state', () => {
    let state = {
        "Active Year": "2024",
        "Current Queue Phase": "VACATION_RANDOM",
        "Current Queue Cycle": 1,
        "Current Serpentine Direction": "FORWARD",
        "Current Queue Cursor": 1,
        "Current Queue Skip State": {},
        "Current Active Window": [
            { turnId: 'turn-1', participantId: 'p1', cycle: 1, direction: 'FORWARD', position: 0 }
        ],
        "Current Directional Window": [],
        "Current Directional Window Completed": [],
        "Active Window Generation": 5
    };
    const roster = [
        { participantId: 'p1', disposition: 'ELIGIBLE' },
        { participantId: 'p2', disposition: 'ELIGIBLE' }
    ];

    let result = calculateNextQueueState(state, roster, {
        action: 'COMPLETE', completedTurnId: 'wrong-id', windowSize: 1, movementMode: 'SERPENTINE'
    });

    assert.equal(result.nextState, state); // exact same object returned unmodified
    assert.equal(result.nextState["Active Window Generation"], 5);
});

test('processQueueMutation: Rejects wrong participant', () => {
    resetMock();
    initializeOrUpdateWorkbook();

    // Seed config
    writeConfigState({
        "Current Active Window": JSON.stringify([{ turnId: 'turn1', participantId: 'uuid1' }])
    });

    // Submit with wrong participant ID
    const res = processQueueMutation('uuid2', 'turn1', {}, () => []);
    assert.equal(res.ok, false);
    assert.match(res.message, /Wrong participant/);
});

test('processQueueMutation: Rejects stale turn', () => {
    resetMock();
    initializeOrUpdateWorkbook();

    writeConfigState({
        "Current Active Window": JSON.stringify([{ turnId: 'turn1', participantId: 'uuid1' }])
    });

    const res = processQueueMutation('uuid1', 'staleTurn', {}, () => []);
    assert.equal(res.ok, false);
    assert.match(res.message, /no longer active/);
});

test('calculateNextQueueState: Dispositions handled correctly, dynamic size', () => {
    let state = {
        "Active Year": "2024",
        "Current Queue Phase": "VACATION_RANDOM",
        "Current Queue Cycle": 1,
        "Current Serpentine Direction": "FORWARD",
        "Current Queue Cursor": 0,
        "Current Queue Skip State": {},
        "Current Active Window": [],
        "Current Directional Window": [],
        "Current Directional Window Completed": [],
        "Active Window Generation": 1
    };
    const roster = [
        { participantId: 'p1', disposition: 'EXCLUDED' },
        { participantId: 'p2', disposition: 'ELIGIBLE' },
        { participantId: 'p3', disposition: 'SKIP_ONCE' },
        { participantId: 'p4', disposition: 'ELIGIBLE' },
        { participantId: 'p5', disposition: 'PERMANENT_PASS' }
    ];

    // Set pending skip for p3
    state["Current Queue Skip State"] = { 'p3': 1 };

    let result = calculateNextQueueState(state, roster, {
        action: 'RECONCILE', windowSize: 2, movementMode: 'SERPENTINE'
    });

    assert.equal(result.nextState["Current Active Window"].length, 2);
    // p1 excluded, p2 eligible.
    // p3 has SKIP_ONCE and a count of 1. It consumes it and continues.
    // p4 is eligible.
    assert.equal(result.nextState["Current Active Window"][0].participantId, 'p2');
    assert.equal(result.nextState["Current Active Window"][1].participantId, 'p4');

    // p3's skip count should be consumed (deleted since it went to 0)
    assert.equal(result.nextState["Current Queue Skip State"]['p3'], undefined);
});
