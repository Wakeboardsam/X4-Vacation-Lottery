const { apiResponse, getHeaderMap, findRowIndex, getDuplicates } = require('../Utils.gs');
const assert = require('node:assert/strict');
const { test } = require('node:test');
const { resetMock } = require('./gas-mock');

test('Utils: getHeaderMap', () => {
    const data = [['Name', 'Age', '  Role '], ['Alice', 30, 'Admin']];
    const map = getHeaderMap(data);
    assert.equal(map['Name'], 0);
    assert.equal(map['Age'], 1);
    assert.equal(map['Role'], 2); // Should be trimmed
});

test('Utils: findRowIndex', () => {
     const data = [['Key', 'Value'], ['Schema Version', '1'], ['Phase', 'SETUP']];
     assert.equal(findRowIndex(data, 0, 'Phase'), 2);
     assert.equal(findRowIndex(data, 0, 'schema version'), 1); // case insensitive
     assert.equal(findRowIndex(data, 0, 'Missing'), -1);
});
