const REQUIRED_SHEETS = [
  "Config",
  "Turn Management",
  "Week Availability",
  "Admin Options",
  "Rules & Tips",
  "Weekend Coverage",
  "Holiday Coverage",
  "Soft Holiday Warnings",
  "Transfer Offers",
  "Transfer History"
];

const REQUIRED_HEADERS = {
  "Config": ["Key", "Value", "Description"],
  "Turn Management": [
    "Name", "PIN", "Phone Number", "Seniority Position", "Lottery Position",
    "Active for Year", "Vacation Phase Enabled", "Vacation Week Target Override",
    "Weekend Phase Enabled", "Weekend Assignment Maximum", "Holiday Volunteer",
    "Mandatory Holiday Eligible", "Transfer Giver", "Transfer Receiver",
    "Had Spring Break Last Year", "Had Christmas Week Last Year",
    "Worked Any Official Holiday Last Year", "Rules Acknowledged Year"
  ],
  "Week Availability": ["Week ID", "Week Start", "Week End", "Prime Classification", "Special Week", "Capacity Override"],
  "Admin Options": ["Option", "Value", "Description", "Sensitive"],
  "Rules & Tips": ["Rule ID", "Context", "Title", "Content", "Display Order", "Enabled"],
  "Weekend Coverage": ["Weekend ID", "Saturday Date", "Saturday First Call", "Sunday Date", "Sunday First Call"],
  "Holiday Coverage": ["Holiday ID", "Holiday Name", "Observed Date", "Call 2 Assignee", "Call 1 Assignee"],
  "Soft Holiday Warnings": ["Warning ID", "Name", "Date", "Enabled"],
  "Transfer Offers": ["Offer ID", "Assignment Type", "Assignment ID", "Position", "Original Assignee", "Status", "Offered At", "Accepted By", "Accepted At"],
  "Transfer History": ["Transfer ID", "Offer ID", "Assignment Type", "Assignment ID", "Position", "Original Assignee", "New Assignee", "Transferred At"]
};

// Requires a specific function for Utils.gs dependencies since it runs in the same GAS environment
const _getHMap = typeof getHeaderMap === 'function' ? getHeaderMap : (typeof global !== 'undefined' && global.getHeaderMap ? global.getHeaderMap : (require('./Utils.gs').getHeaderMap));
const _findRowIdx = typeof findRowIndex === 'function' ? findRowIndex : (typeof global !== 'undefined' && global.findRowIndex ? global.findRowIndex : (require('./Utils.gs').findRowIndex));
const _getDups = typeof getDuplicates === 'function' ? getDuplicates : (typeof global !== 'undefined' && global.getDuplicates ? global.getDuplicates : (require('./Utils.gs').getDuplicates));
const _DEFAULT_CONFIG = typeof DEFAULT_CONFIG !== 'undefined' ? DEFAULT_CONFIG : (typeof global !== 'undefined' && global.DEFAULT_CONFIG ? global.DEFAULT_CONFIG : (require('./State.gs').DEFAULT_CONFIG));

/**
 * Ensures workbook is initialized correctly based on schema
 */
function initializeOrUpdateWorkbook() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, message: 'Could not obtain lock for schema update. Please try again later.' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. INSPECT AND PLAN
    const plan = {
      sheetsToAdd: [],
      headersToAppend: {},
      configKeysToAdd: {},
      validationsToApply: [],
      warnings: [],
      conflicts: []
    };

    let hasChanges = false;

    // Validate sheets and headers
    for (const sheetName of REQUIRED_SHEETS) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        plan.sheetsToAdd.push(sheetName);
        plan.headersToAppend[sheetName] = REQUIRED_HEADERS[sheetName];
        hasChanges = true;
      } else {
        const data = sheet.getDataRange().getValues();
        let headers = [];

        if (data.length > 0) {
          headers = data[0].map(h => String(h || '').trim());

          // Check for identifiable header row (must contain at least one string value if populated)
          const hasAnyHeader = headers.some(h => h.length > 0);
          if (!hasAnyHeader && data.length > 1) {
            plan.conflicts.push(`Sheet '${sheetName}' is populated but lacks an identifiable header row in row 1.`);
            continue; // Can't process safely
          }
        }

        const dups = _getDups(headers.filter(h => h.length > 0));
        if (dups.length > 0) {
          plan.conflicts.push(`Duplicate headers found in sheet '${sheetName}': ${dups.join(', ')}`);
          continue;
        }

        const existingHeadersSet = new Set(headers);
        const missingHeaders = REQUIRED_HEADERS[sheetName].filter(h => !existingHeadersSet.has(h));

        if (missingHeaders.length > 0) {
          plan.headersToAppend[sheetName] = missingHeaders;
          hasChanges = true;
        }

        // Plan validations (only apply to data rows, starting at row 2)
        if (sheetName === 'Turn Management') {
            const checkboxCols = ["Active for Year", "Vacation Phase Enabled", "Weekend Phase Enabled",
                                  "Holiday Volunteer", "Mandatory Holiday Eligible", "Transfer Giver",
                                  "Transfer Receiver", "Had Spring Break Last Year",
                                  "Had Christmas Week Last Year", "Worked Any Official Holiday Last Year"];

            for (const col of checkboxCols) {
                const colIdx = headers.indexOf(col);
                if (colIdx > -1) {
                    // Check for existing non-boolean data
                    for (let r = 1; r < data.length; r++) {
                        const val = data[r][colIdx];
                        if (val !== '' && val !== true && val !== false && String(val).toLowerCase() !== 'true' && String(val).toLowerCase() !== 'false') {
                             plan.conflicts.push(`Existing non-Boolean value found in checkbox column '${col}' of 'Turn Management' at row ${r+1}`);
                        }
                    }
                    plan.validationsToApply.push({ sheetName, col: colIdx + 1, type: 'CHECKBOX' });
                }
            }
        } else if (sheetName === 'Admin Options') {
            const colIdx = headers.indexOf('Sensitive');
            if (colIdx > -1) {
                 for (let r = 1; r < data.length; r++) {
                    const val = data[r][colIdx];
                    if (val !== '' && val !== true && val !== false && String(val).toLowerCase() !== 'true' && String(val).toLowerCase() !== 'false') {
                         plan.conflicts.push(`Existing non-Boolean value found in checkbox column 'Sensitive' of 'Admin Options' at row ${r+1}`);
                    }
                }
                plan.validationsToApply.push({ sheetName, col: colIdx + 1, type: 'CHECKBOX' });
            }
        } else if (sheetName === 'Rules & Tips' || sheetName === 'Soft Holiday Warnings') {
             const colIdx = headers.indexOf('Enabled');
             if (colIdx > -1) {
                 for (let r = 1; r < data.length; r++) {
                    const val = data[r][colIdx];
                    if (val !== '' && val !== true && val !== false && String(val).toLowerCase() !== 'true' && String(val).toLowerCase() !== 'false') {
                         plan.conflicts.push(`Existing non-Boolean value found in checkbox column 'Enabled' of '${sheetName}' at row ${r+1}`);
                    }
                }
                plan.validationsToApply.push({ sheetName, col: colIdx + 1, type: 'CHECKBOX' });
             }
        }
      }
    }

    // Check Config state
    const configSheet = ss.getSheetByName('Config');
    if (configSheet) {
      const data = configSheet.getDataRange().getValues();
      const map = _getHMap(data);
      const keyCol = map['Key'];

      if (keyCol !== undefined) {
        const existingKeys = new Set();
        const dups = new Set();
        for (let i = 1; i < data.length; i++) {
          const k = String(data[i][keyCol] || '').trim();
          if (k) {
             if (existingKeys.has(k)) dups.add(k);
             existingKeys.add(k);
          }
        }

        if (dups.size > 0) {
           plan.conflicts.push(`Duplicate Config keys found: ${Array.from(dups).join(', ')}`);
        } else {
           for (const k in _DEFAULT_CONFIG) {
              if (!existingKeys.has(k)) {
                  plan.configKeysToAdd[k] = _DEFAULT_CONFIG[k];
                  hasChanges = true;
              }
           }
        }
      }
    } else {
        // Handled by sheetsToAdd logic which will append missing headers and then we can append the keys
        for (const k in _DEFAULT_CONFIG) {
           plan.configKeysToAdd[k] = _DEFAULT_CONFIG[k];
        }
        hasChanges = true;
    }

    // IF CONFLICTS, ABORT
    if (plan.conflicts.length > 0) {
       return {
           ok: false,
           data: { report: plan, changed: false },
           message: 'Initialization aborted due to schema or data conflicts.'
       };
    }

    // IF NO CHANGES REQUIRED, RETURN EARLY
    // For idempotency, we return changed: false if only validations were applied (since we can't easily read existing validations to see if they were already there, we always re-apply them but don't count it as a "schema change").
    // We only return early if nothing changed AND no validations to apply. Wait, if we return early, validations won't be applied. Let's not return early if we have validations to apply.
    if (!hasChanges) {
       // Apply validations and formatting even if "no changes" to headers/sheets/config
       for (const v of plan.validationsToApply) {
           const sheet = ss.getSheetByName(v.sheetName);
           const lastRow = Math.max(sheet.getLastRow(), 2);
           const range = sheet.getRange(2, v.col, lastRow - 1, 1);
           if (v.type === 'CHECKBOX') {
               const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
               range.setDataValidation(rule);
           }
       }
       const tmSheet = ss.getSheetByName('Turn Management');
       if (tmSheet) {
           const tmM = _getHMap(tmSheet.getDataRange().getValues());
           if (tmM['PIN'] !== undefined) {
                tmSheet.getRange(2, tmM['PIN'] + 1, Math.max(tmSheet.getLastRow() - 1, 1)).setNumberFormat('@');
           }
           if (tmM['Phone Number'] !== undefined) {
                tmSheet.getRange(2, tmM['Phone Number'] + 1, Math.max(tmSheet.getLastRow() - 1, 1)).setNumberFormat('@');
           }
       }

       return {
           ok: true,
           data: { report: plan, changed: false },
           message: 'Workbook is already fully initialized.'
       };
    }

    // 2. APPLY CHANGES

    // Add sheets
    for (const sheetName of plan.sheetsToAdd) {
        ss.insertSheet(sheetName);
    }

    // Append headers
    for (const sheetName in plan.headersToAppend) {
        const sheet = ss.getSheetByName(sheetName);
        const missing = plan.headersToAppend[sheetName];
        if (missing.length > 0) {
            const numCols = sheet.getLastColumn();
            if (numCols === 0) {
                // Empty sheet
                sheet.getRange(1, 1, 1, missing.length).setValues([missing]);
            } else {
                // Append to existing
                sheet.getRange(1, numCols + 1, 1, missing.length).setValues([missing]);
            }
        }
    }

    // Inject missing config keys
    const configSh = ss.getSheetByName('Config');
    const cData = configSh.getDataRange().getValues();
    const cMap = _getHMap(cData);
    const cKeyCol = cMap['Key'];
    const cValCol = cMap['Value'];

    if (cKeyCol !== undefined && cValCol !== undefined) {
        let lastRow = configSh.getLastRow();
        const keysToAdd = Object.keys(plan.configKeysToAdd);
        if (keysToAdd.length > 0) {
           const rowsToAppend = keysToAdd.map(k => {
               const row = new Array(configSh.getLastColumn()).fill('');
               row[cKeyCol] = k;
               row[cValCol] = plan.configKeysToAdd[k];
               return row;
           });
           configSh.getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
        }
    }

    // Apply validations
    for (const v of plan.validationsToApply) {
        const sheet = ss.getSheetByName(v.sheetName);
        const lastRow = Math.max(sheet.getLastRow(), 2); // ensure we apply validation even if empty
        const range = sheet.getRange(2, v.col, lastRow - 1, 1);

        if (v.type === 'CHECKBOX') {
            const rule = SpreadsheetApp.newDataValidation().requireCheckbox().build();
            range.setDataValidation(rule);
        }
        // Could expand this to numeric, list validations later based on prompt details,
        // but prompt explicitly calls out checkboxes for all boolean participant controls.
    }

    // Enforce Turn Management text formats for PIN and Phone Number
    const tmSheet = ss.getSheetByName('Turn Management');
    if (tmSheet) {
        const tmM = _getHMap(tmSheet.getDataRange().getValues());
        if (tmM['PIN'] !== undefined) {
             tmSheet.getRange(2, tmM['PIN'] + 1, Math.max(tmSheet.getLastRow() - 1, 1)).setNumberFormat('@'); // Text format to preserve zeroes
        }
        if (tmM['Phone Number'] !== undefined) {
             tmSheet.getRange(2, tmM['Phone Number'] + 1, Math.max(tmSheet.getLastRow() - 1, 1)).setNumberFormat('@');
        }
    }


    return {
        ok: true,
        data: { report: plan, changed: true },
        message: 'Workbook schema initialized and updated successfully.'
    };

  } catch (e) {
      return { ok: false, message: `Unexpected error during initialization: ${e.message}` };
  } finally {
      lock.releaseLock();
  }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        initializeOrUpdateWorkbook,
        REQUIRED_SHEETS,
        REQUIRED_HEADERS
    };
}
