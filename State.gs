// Required initial keys and fresh-workbook values
const DEFAULT_CONFIG = {
  "Schema Version": "1",
  "Active Year": "",
  "Setup State": "SETUP_EMPTY",
  "Current Phase": "SETUP",
  "Phase Ready State": "NONE",
  "Current Vacation Round": "0",
  "Current Serpentine Direction": "FORWARD",
  "Current Queue Phase": "",
  "Current Queue Order Source": "",
  "Current Queue Cursor": "",
  "Current Queue Cycle": "0",
  "Current Active Window": "[]",
  "Current Directional Window": "[]",
  "Current Directional Window Completed": "[]",
  "Active Window Generation": "0"
};

const VALID_PHASES = [
  "SETUP",
  "VACATION_SENIORITY",
  "VACATION_RANDOM",
  "WEEKEND",
  "HOLIDAY_VOLUNTEER",
  "HOLIDAY_MANDATORY",
  "TRANSFER_OFFER_COLLECTION",
  "TRANSFER_RECEIVER",
  "COMPLETE"
];

const VALID_READY_STATES = [
  "NONE",
  "READY_VACATION_SENIORITY",
  "READY_VACATION_RANDOM",
  "READY_WEEKEND",
  "READY_HOLIDAY_VOLUNTEER",
  "READY_HOLIDAY_MANDATORY",
  "READY_TRANSFER_OFFER_COLLECTION",
  "READY_TRANSFER_RECEIVER",
  "READY_COMPLETE"
];

const VALID_SETUP_STATES = [
  "SETUP_EMPTY",
  "SETUP_AUTO_FILLED",
  "SETUP_REVIEW",
  "SETUP_CONFIRMED"
];

const VALID_SERPENTINE_DIRECTIONS = [
  "FORWARD",
  "BACKWARD"
];


/**
 * Reads all Config values from the spreadsheet into an object mapping keys to values.
 * @returns {object} Config map
 */
function readConfigState_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Config');
  if (!sheet) {
    throw new Error('Config sheet is missing');
  }

  const data = sheet.getDataRange().getValues();
  if (data.length === 0) {
    throw new Error('Config sheet is empty');
  }

  const map = typeof getHeaderMap_ === 'function' ? getHeaderMap_(data) : (typeof global !== 'undefined' && global.getHeaderMap_ ? global.getHeaderMap_(data) : (require('./Utils.gs').getHeaderMap(data)));
  if (map['Key'] === undefined || map['Value'] === undefined) {
    throw new Error('Config sheet missing required Key or Value headers');
  }

  const keyCol = map['Key'];
  const valCol = map['Value'];

  const configMap = {};
  for (let i = 1; i < data.length; i++) {
    const k = String(data[i][keyCol] || '').trim();
    if (k) {
      if (configMap[k] !== undefined) {
          throw new Error(`Duplicate Config key found: ${k}`);
      }
      configMap[k] = data[i][valCol];
    }
  }

  return configMap;
}

/**
 * Validates a single configuration key-value pair against allowed states.
 * @param {string} key
 * @param {any} val
 * @throws {Error} if validation fails
 */
function validateConfigValue_(key, val) {
  const strVal = String(val || '').trim();

  if (key === 'Setup State' && !VALID_SETUP_STATES.includes(strVal) && strVal !== '') {
     throw new Error(`Invalid Setup State: ${strVal}`);
  }
  if (key === 'Current Phase' && !VALID_PHASES.includes(strVal)) {
     throw new Error(`Invalid Current Phase: ${strVal}`);
  }
  if (key === 'Phase Ready State' && !VALID_READY_STATES.includes(strVal)) {
     throw new Error(`Invalid Phase Ready State: ${strVal}`);
  }
  if (key === 'Current Serpentine Direction' && !VALID_SERPENTINE_DIRECTIONS.includes(strVal) && strVal !== '') {
     throw new Error(`Invalid Serpentine Direction: ${strVal}`);
  }
  if (['Current Active Window', 'Current Directional Window', 'Current Directional Window Completed'].includes(key)) {
     try {
       const parsed = JSON.parse(val || '[]');
       if (!Array.isArray(parsed)) throw new Error();
     } catch (e) {
       throw new Error(`Invalid queue-list serialization for ${key}: ${val}`);
     }
  }

  if (['Current Vacation Round', 'Current Queue Cycle', 'Active Window Generation'].includes(key)) {
      if (strVal !== '' && isNaN(Number(strVal))) {
          throw new Error(`Invalid numeric value for ${key}: ${val}`);
      }
  }
}

/**
 * Validates an entire config map.
 * @param {object} configMap
 */
function validateConfigState_(configMap) {
   for (const key in DEFAULT_CONFIG) {
       if (configMap[key] === undefined) {
           throw new Error(`Missing required Config key: ${key}`);
       }
       validateConfigValue_(key, configMap[key]);
   }
}

/**
 * Safely writes a specific subset of state values to Config.
 * @param {object} updates Map of keys to new values
 */
function writeConfigState_(updates) {
    const lock = LockService.getScriptLock();
    // Wait up to 10 seconds for other processes to finish
    if (!lock.tryLock(10000)) {
        throw new Error('Could not obtain lock to update Config state');
    }

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Config');
        if (!sheet) throw new Error('Config sheet missing');

        const dataRange = sheet.getDataRange();
        const data = dataRange.getValues();
        const formulas = dataRange.getFormulas();

        const map = typeof getHeaderMap_ === 'function' ? getHeaderMap_(data) : (typeof global !== 'undefined' && global.getHeaderMap_ ? global.getHeaderMap_(data) : (require('./Utils.gs').getHeaderMap(data)));
        if (map['Key'] === undefined || map['Value'] === undefined) {
            throw new Error('Config sheet missing required Key or Value headers');
        }

        const keyCol = map['Key'];
        const valCol = map['Value'];

        // Validate complete proposed state
        const completeMap = {};
        for (let i = 1; i < data.length; i++) {
            const k = String(data[i][keyCol] || '').trim();
            if (k) {
                completeMap[k] = data[i][valCol];
            }
        }

        const proposedState = { ...completeMap };
        for (const k in updates) {
            if (DEFAULT_CONFIG[k] === undefined) {
                throw new Error(`Attempting to write unknown Config key: ${k}`);
            }
            proposedState[k] = updates[k];
        }
        validateConfigState_(proposedState);

        // Fetch just the Value column range to update
        // We only modify the in-memory array of the Value column
        const maxRow = data.length;
        const valColRange = sheet.getRange(1, valCol + 1, maxRow, 1);
        const valColValues = valColRange.getValues();
        const valColFormulas = valColRange.getFormulas();

        for (const k in updates) {
            let rowIndex = typeof findRowIndex_ === 'function' ? findRowIndex_(data, keyCol, k) : (typeof global !== 'undefined' && global.findRowIndex_ ? global.findRowIndex_(data, keyCol, k) : (require('./Utils.gs').findRowIndex(data, keyCol, k)));

            if (rowIndex === -1) {
                 throw new Error(`Config key ${k} not found in sheet`);
            }

            if (valColFormulas[rowIndex][0] !== '') {
                 throw new Error(`Conflict: Config key ${k} contains a formula and cannot be overwritten`);
            }

            valColValues[rowIndex][0] = String(updates[k]);
        }

        // Commit complete affected Value range with one batch operation
        valColRange.setValues(valColValues);
    } finally {
        lock.releaseLock();
    }
}

// Node.js module export for testing
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    DEFAULT_CONFIG,
    VALID_PHASES,
    VALID_SETUP_STATES,
    VALID_READY_STATES,
    VALID_SERPENTINE_DIRECTIONS,
    readConfigState: readConfigState_,
    validateConfigState: validateConfigState_,
    validateConfigValue: validateConfigValue_,
    writeConfigState: writeConfigState_
  };
}
