/**
 * Standardizes API responses
 * @param {boolean} ok
 * @param {any} data
 * @param {string} message
 * @returns {object}
 */
function apiResponse(ok, data = null, message = '') {
  return { ok, data, message };
}

/**
 * Creates a map of header names to their 0-indexed column positions.
 * @param {any[][]} data 2D array of sheet data
 * @returns {object} map of string header name to integer column index
 */
function getHeaderMap(data) {
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
function findRowIndex(data, colIndex, targetValue) {
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
function getDuplicates(arr) {
  const seen = new Set();
  const dupes = new Set();
  for (const item of arr) {
    if (seen.has(item)) dupes.add(item);
    seen.add(item);
  }
  return Array.from(dupes);
}

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    apiResponse,
    getHeaderMap,
    findRowIndex,
    getDuplicates
  };
}
