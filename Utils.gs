/**
 * Standardizes API responses
 * @param {boolean} ok
 * @param {any} data
 * @param {string} message
 * @returns {object}
 */
function apiResponse_(ok, data = null, message = '') {
  return { ok, data, message };
}

/**
 * Creates a map of header names to their 0-indexed column positions.
 * @param {any[][]} data 2D array of sheet data
 * @returns {object} map of string header name to integer column index
 */
function getHeaderMap_(data) {
  if (!data || data.length === 0) return {};
  const headers = data[0];
  const map = {};
  for (let i = 0; i < headers.length; i++) {
    const h = String(headers[i] || '').trim();
    if (h) {
      map[h] = i;
    }
  }
  return map;
}

/**
 * Finds a row index based on a specific column's value (case-insensitive for text).
 * @param {any[][]} data
 * @param {number} colIndex
 * @param {string|number} targetValue
 * @returns {number} 0-indexed row number, or -1 if not found
 */
function findRowIndex_(data, colIndex, targetValue) {
  if (!data || data.length === 0 || colIndex < 0 || targetValue === null || targetValue === undefined) {
    return -1;
  }

  const searchStr = String(targetValue).trim().toLowerCase();

  // start from row 1 to skip header
  for (let i = 1; i < data.length; i++) {
    const valStr = String(data[i][colIndex] || '').trim().toLowerCase();
    if (valStr === searchStr) {
      return i;
    }
  }
  return -1;
}

/**
 * Returns all duplicates in an array
 * @param {any[]} arr
 * @returns {any[]} duplicates
 */
function getDuplicates_(arr) {
  const seen = new Set();
  const dupes = new Set();
  for (const item of arr) {
    if (seen.has(item)) dupes.add(item);
    seen.add(item);
  }
  return Array.from(dupes);
}

/**
 * Throws a schema error when required headers are absent.
 * @param {object} map Header map
 * @param {string[]} required Required header names
 * @param {string} sheetName Display name for the error
 */
function requireHeaders_(map, required, sheetName) {
  const missing = required.filter(header => map[header] === undefined);
  if (missing.length > 0) {
    const err = new Error(`Schema conflict in ${sheetName}: missing required header(s): ${missing.join(', ')}`);
    err.code = 'SCHEMA_CONFLICT';
    throw err;
  }
}

/**
 * Validates stable participant identities for every populated Turn Management row.
 * PIN remains an authentication credential and is never accepted as identity.
 * @param {any[][]} data Turn Management matrix
 * @param {object} map Header map
 */
function validateParticipantIds_(data, map) {
  requireHeaders_(map, ['Name', 'Participant ID', 'PIN'], 'Turn Management');

  const blankNames = [];
  const namesById = {};
  const duplicateNames = {};

  for (let r = 1; r < data.length; r++) {
    const name = String(data[r][map['Name']] || '').trim();
    const participantId = String(data[r][map['Participant ID']] || '').trim();
    const pin = String(data[r][map['PIN']] || '').trim();
    const populated = name !== '' || participantId !== '' || pin !== '';
    if (!populated) continue;

    const displayName = name || '(unnamed participant)';
    if (!participantId) {
      blankNames.push(displayName);
      continue;
    }

    if (namesById[participantId] !== undefined) {
      if (!duplicateNames[participantId]) duplicateNames[participantId] = [namesById[participantId]];
      duplicateNames[participantId].push(displayName);
    } else {
      namesById[participantId] = displayName;
    }
  }

  const details = [];
  if (blankNames.length > 0) details.push(`blank Participant ID: ${blankNames.join(', ')}`);
  for (const participantId in duplicateNames) {
    details.push(`duplicate Participant ID '${participantId}': ${duplicateNames[participantId].join(', ')}`);
  }

  if (details.length > 0) {
    const err = new Error(`Participant roster identity conflict (${details.join('; ')}). Correct Turn Management before continuing.`);
    err.code = 'PARTICIPANT_ID_CONFLICT';
    err.participantMessage = 'Vacation selection is temporarily unavailable because the participant roster requires administrator correction. No changes were made.';
    throw err;
  }
}

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    apiResponse: apiResponse_,
    getHeaderMap: getHeaderMap_,
    findRowIndex: findRowIndex_,
    getDuplicates: getDuplicates_,
    requireHeaders: requireHeaders_,
    validateParticipantIds: validateParticipantIds_
  };
}
