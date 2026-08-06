const { validateConfigValue } = require('../State.gs');
const assert = require('node:assert/strict');
const { test } = require('node:test');

test('State: Phase and Phase Ready State are independent', () => {
    // Phase and Phase Ready State don't mutually constrain each other in validation
    assert.doesNotThrow(() => {
         validateConfigValue('Current Phase', 'SETUP');
         validateConfigValue('Phase Ready State', 'VACATION_SENIORITY'); // valid independently
    });
});

test('State: Queue list serialization and generation check', () => {
    // Valid
    assert.doesNotThrow(() => validateConfigValue('Current Active Window', '["A", "B"]'));
    assert.doesNotThrow(() => validateConfigValue('Active Window Generation', '5'));

    // Invalid
    assert.throws(() => validateConfigValue('Current Active Window', 'invalid json'), /Invalid queue-list serialization/);
    assert.throws(() => validateConfigValue('Active Window Generation', 'NaN'), /Invalid numeric value/);
});

test('State: Invalid direction, round rejected', () => {
    assert.throws(() => validateConfigValue('Current Serpentine Direction', 'UP'), /Invalid Serpentine Direction/);
    assert.throws(() => validateConfigValue('Current Vacation Round', 'One'), /Invalid numeric value/);
});
