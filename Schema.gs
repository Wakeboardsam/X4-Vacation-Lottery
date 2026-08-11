// Schema Service
const REQUIRED_SHEETS = [
  'Turn Management',
  'Week Availability',
  'Admin Options',
  'Rules & Tips',
  'Soft Holiday Warnings',
  'Config'
];

const REQUIRED_HEADERS = {
  'Turn Management': [
    "Name", "Participant ID", "PIN", "Phone Number", "Active for Year", "Seniority Position", "Lottery Position",
    "Vacation Phase Enabled", "Vacation Week Target Override", "Weekend Phase Enabled",
    "Weekend Assignment Maximum", "Holiday Volunteer", "Mandatory Holiday Eligible",
    "Transfer Giver", "Transfer Receiver", "Had Spring Break Last Year", "Had Christmas Week Last Year",
    "Worked Any Official Holiday Last Year", "Week Availability Capacity Override", "Rules Acknowledged Year"
  ],
  'Week Availability': [
    "Week Start Date", "Prime Classification", "Special Week", "Capacity Override",
    "Person1", "Person2", "Person3", "Person4"
  ],
  'Admin Options': ["Setting", "Value", "Sensitive", "Description"],
  'Rules & Tips': ["Display Order", "Rule Text", "Enabled"],
  'Soft Holiday Warnings': ["Warning Text", "Enabled"],
  'Config': ["Key", "Value", "Description"]
};

const _apiResponse = typeof apiResponse_ === 'function' ? apiResponse_ : (typeof global !== 'undefined' && global.apiResponse_ ? global.apiResponse_ : (require('./Utils.gs').apiResponse));
const _getHMap = typeof getHeaderMap_ === 'function' ? getHeaderMap_ : (typeof global !== 'undefined' && global.getHeaderMap_ ? global.getHeaderMap_ : (require('./Utils.gs').getHeaderMap));
const _getDups = typeof getDuplicates_ === 'function' ? getDuplicates_ : (typeof global !== 'undefined' && global.getDuplicates_ ? global.getDuplicates_ : (require('./Utils.gs').getDuplicates));
const _DEFAULT_CONFIG_REF = typeof DEFAULT_CONFIG !== 'undefined' ? DEFAULT_CONFIG : (typeof global !== 'undefined' && global.DEFAULT_CONFIG ? global.DEFAULT_CONFIG : (require('./State.gs').DEFAULT_CONFIG));
const _VALID_PHASES = typeof VALID_PHASES !== 'undefined' ? VALID_PHASES : (typeof global !== 'undefined' && global.VALID_PHASES ? global.VALID_PHASES : (require('./State.gs').VALID_PHASES));
const _VALID_SETUP_STATES = typeof VALID_SETUP_STATES !== 'undefined' ? VALID_SETUP_STATES : (typeof global !== 'undefined' && global.VALID_SETUP_STATES ? global.VALID_SETUP_STATES : (require('./State.gs').VALID_SETUP_STATES));
const _VALID_SERPENTINE_DIRECTIONS = typeof VALID_SERPENTINE_DIRECTIONS !== 'undefined' ? VALID_SERPENTINE_DIRECTIONS : (typeof global !== 'undefined' && global.VALID_SERPENTINE_DIRECTIONS ? global.VALID_SERPENTINE_DIRECTIONS : (require('./State.gs').VALID_SERPENTINE_DIRECTIONS));
const _VALID_READY_STATES = typeof VALID_READY_STATES !== 'undefined' ? VALID_READY_STATES : (typeof global !== 'undefined' && global.VALID_READY_STATES ? global.VALID_READY_STATES : (require('./State.gs').VALID_READY_STATES));


function planSchema_(ss) {
    const plan = {
        sheetsToAdd: [],
        headersToAppend: {},
        configKeysToAdd: {},
        validationsToApply: [],
        formatsToApply: [],
        conflicts: [],
        warnings: [],
        hasChanges: false
    };

    // Check Sheets
    for (const sheetName of REQUIRED_SHEETS) {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            plan.sheetsToAdd.push(sheetName);
            plan.headersToAppend[sheetName] = REQUIRED_HEADERS[sheetName];

            // Queue up validations for the new sheet
            if (sheetName === 'Turn Management') {
                const headMap = {};
                REQUIRED_HEADERS[sheetName].forEach((h, i) => headMap[h] = i);

                const checkboxes = ["Active for Year", "Vacation Phase Enabled", "Weekend Phase Enabled", "Holiday Volunteer", "Mandatory Holiday Eligible", "Transfer Giver", "Transfer Receiver", "Had Spring Break Last Year", "Had Christmas Week Last Year", "Worked Any Official Holiday Last Year"];
                for (const col of checkboxes) plan.validationsToApply.push({ sheetName, col: headMap[col] + 1, type: 'CHECKBOX' });

                const posInts = ["Seniority Position", "Lottery Position", "Vacation Week Target Override", "Weekend Assignment Maximum", "Week Availability Capacity Override"];
                for (const col of posInts) plan.validationsToApply.push({ sheetName, col: headMap[col] + 1, type: 'POS_INT' });

                plan.validationsToApply.push({ sheetName, col: headMap['Rules Acknowledged Year'] + 1, type: 'YEAR' });
                plan.formatsToApply.push({ sheetName, col: headMap['PIN'] + 1, format: '@' });
                plan.formatsToApply.push({ sheetName, col: headMap['Phone Number'] + 1, format: '@' });
            } else if (sheetName === 'Week Availability') {
                const headMap = {};
                REQUIRED_HEADERS[sheetName].forEach((h, i) => headMap[h] = i);
                plan.validationsToApply.push({ sheetName, col: headMap['Capacity Override'] + 1, type: 'STRICT_POS_INT' });
                plan.validationsToApply.push({ sheetName, col: headMap['Prime Classification'] + 1, type: 'LIST_PRIME' });
                plan.validationsToApply.push({ sheetName, col: headMap['Special Week'] + 1, type: 'LIST_SPECIAL' });
            } else if (sheetName === 'Admin Options') {
                plan.validationsToApply.push({ sheetName, col: 3, type: 'CHECKBOX' }); // Sensitive
            } else if (sheetName === 'Rules & Tips') {
                plan.validationsToApply.push({ sheetName, col: 3, type: 'CHECKBOX' }); // Enabled
            } else if (sheetName === 'Soft Holiday Warnings') {
                plan.validationsToApply.push({ sheetName, col: 2, type: 'CHECKBOX' }); // Enabled
            }
        } else {
            const data = sheet.getDataRange().getValues();
            let headers = data.length > 0 ? data[0] : [];
            let hasAnyHeader = headers.some(h => String(h).trim().length > 0);

            if (!hasAnyHeader && data.length > 1) {
                plan.conflicts.push(`Sheet '${sheetName}' is populated but lacks an identifiable header row in row 1.`);
                continue;
            }

            const dups = _getDups(headers.filter(h => String(h).trim().length > 0));
            if (dups.length > 0) {
                plan.conflicts.push(`Duplicate headers found in sheet '${sheetName}': ${dups.join(', ')}`);
                continue;
            }

            const existingHeadersSet = new Set(headers.map(h => String(h).trim()));
            const missingHeaders = REQUIRED_HEADERS[sheetName].filter(h => !existingHeadersSet.has(h));

            if (missingHeaders.length > 0) {
                plan.headersToAppend[sheetName] = missingHeaders;

            }

            const headMap = _getHMap(data);

            if (sheetName === 'Week Availability') {
                if (headMap['Capacity Override'] !== undefined) {
                    const col = 'Capacity Override';
                    for (let r = 1; r < data.length; r++) {
                        const val = data[r][headMap[col]];
                        if (val !== '' && (!Number.isInteger(Number(val)) || Number(val) <= 0)) {
                            plan.conflicts.push(`Existing invalid strictly positive integer found in column '${col}' at row ${r+1}`);
                        }
                    }
                    const rule = buildValidationRule_('STRICT_POS_INT');
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    const existingTop = sheet.getRange(2, headMap[col] + 1).getDataValidation();
                    const existingBottom = sheet.getRange(lastRow - 1, headMap[col] + 1).getDataValidation();
                    if (normalizeValidation_(existingTop) !== normalizeValidation_(rule) || normalizeValidation_(existingBottom) !== normalizeValidation_(rule)) {
                        plan.validationsToApply.push({ sheetName, col: headMap[col] + 1, type: 'STRICT_POS_INT' });
                    }
                }
            } else if (sheetName === 'Turn Management') {
                const checkboxes = ["Active for Year", "Vacation Phase Enabled", "Weekend Phase Enabled",
                                  "Holiday Volunteer", "Mandatory Holiday Eligible", "Transfer Giver",
                                  "Transfer Receiver", "Had Spring Break Last Year",
                                  "Had Christmas Week Last Year", "Worked Any Official Holiday Last Year"];

                for (const col of checkboxes) {
                    if (headMap[col] !== undefined) {
                        for (let r = 1; r < data.length; r++) {
                            const val = data[r][headMap[col]];
                            if (val !== '' && val !== true && val !== false && String(val).toLowerCase() !== 'true' && String(val).toLowerCase() !== 'false') {
                                plan.conflicts.push(`Existing non-Boolean value found in checkbox column '${col}' at row ${r+1}`);
                            }
                        }
                        const rule = buildValidationRule_('CHECKBOX');
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    const existingTop = sheet.getRange(2, headMap[col] + 1).getDataValidation();
                    const existingBottom = sheet.getRange(lastRow - 1, headMap[col] + 1).getDataValidation();
                    if (normalizeValidation_(existingTop) !== normalizeValidation_(rule) || normalizeValidation_(existingBottom) !== normalizeValidation_(rule)) {
                        plan.validationsToApply.push({ sheetName, col: headMap[col] + 1, type: 'CHECKBOX' });
                    }
                    }
                }

                // Positive whole number
                const posInts = ["Seniority Position", "Lottery Position", "Vacation Week Target Override", "Weekend Assignment Maximum", "Week Availability Capacity Override"];
                for (const col of posInts) {
                    if (headMap[col] !== undefined) {
                        for (let r = 1; r < data.length; r++) {
                            const val = data[r][headMap[col]];
                            if (val !== '' && (!Number.isInteger(Number(val)) || Number(val) < 0)) {
                                plan.conflicts.push(`Existing invalid positive integer found in column '${col}' at row ${r+1}`);
                            }
                        }
                        const rule = buildValidationRule_('POS_INT');
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    const existingTop = sheet.getRange(2, headMap[col] + 1).getDataValidation();
                    const existingBottom = sheet.getRange(lastRow - 1, headMap[col] + 1).getDataValidation();
                    if (normalizeValidation_(existingTop) !== normalizeValidation_(rule) || normalizeValidation_(existingBottom) !== normalizeValidation_(rule)) {
                        plan.validationsToApply.push({ sheetName, col: headMap[col] + 1, type: 'POS_INT' });
                    }
                    }
                }

                // 4-digit year
                if (headMap['Rules Acknowledged Year'] !== undefined) {
                    const col = 'Rules Acknowledged Year';
                    for (let r = 1; r < data.length; r++) {
                        const val = data[r][headMap[col]];
                        if (val !== '' && !/^\d{4}$/.test(String(val))) {
                             plan.conflicts.push(`Existing invalid year format in column '${col}' at row ${r+1}`);
                        }
                    }
                    const rule = buildValidationRule_('YEAR');
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    const existingTop = sheet.getRange(2, headMap[col] + 1).getDataValidation();
                    const existingBottom = sheet.getRange(lastRow - 1, headMap[col] + 1).getDataValidation();
                    if (normalizeValidation_(existingTop) !== normalizeValidation_(rule) || normalizeValidation_(existingBottom) !== normalizeValidation_(rule)) {
                        plan.validationsToApply.push({ sheetName, col: headMap[col] + 1, type: 'YEAR' });
                    }
                }

                // Text formats
                if (headMap['PIN'] !== undefined) {
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    const fmtTop = sheet.getRange(2, headMap['PIN'] + 1).getNumberFormat();
                    const fmtBottom = sheet.getRange(lastRow - 1, headMap['PIN'] + 1).getNumberFormat();
                    if (fmtTop !== '@' || fmtBottom !== '@')
                     plan.formatsToApply.push({ sheetName, col: headMap['PIN'] + 1, format: '@' });
                }
                if (headMap['Phone Number'] !== undefined) {
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    const fmtTop = sheet.getRange(2, headMap['Phone Number'] + 1).getNumberFormat();
                    const fmtBottom = sheet.getRange(lastRow - 1, headMap['Phone Number'] + 1).getNumberFormat();
                    if (fmtTop !== '@' || fmtBottom !== '@')
                     plan.formatsToApply.push({ sheetName, col: headMap['Phone Number'] + 1, format: '@' });
                }

                // Implicit dropdowns not strictly defined for Module 1 in TM, except "Prime Classification" and "Special Week" mentioned in prompt
                // Prompt: "Dropdown validation for: Prime Classification: "Prime", "Non-Prime", Special Week: "None", "Spring Break", "Christmas""
                // Wait, Prime Classification and Special Week are properties of Vacations, not TM. They might be in a different sheet not explicitly required?
                // The prompt says "Implement and verify: ... Dropdown validation for: Prime Classification..."
                // Since they are not in REQUIRED_HEADERS, we should check if they exist anywhere, or maybe they are in another sheet?
                // If they exist in headers, we apply them.
            } else if (sheetName === 'Admin Options' || sheetName === 'Rules & Tips' || sheetName === 'Soft Holiday Warnings') {
                 const boolCol = sheetName === 'Admin Options' ? 'Sensitive' : 'Enabled';
                 if (headMap[boolCol] !== undefined) {
                     for (let r = 1; r < data.length; r++) {
                        const val = data[r][headMap[boolCol]];
                        if (val !== '' && val !== true && val !== false && String(val).toLowerCase() !== 'true' && String(val).toLowerCase() !== 'false') {
                             plan.conflicts.push(`Existing non-Boolean value found in checkbox column '${boolCol}' of '${sheetName}' at row ${r+1}`);
                        }
                    }

                    const rule = buildValidationRule_('CHECKBOX');
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    const existingTop = sheet.getRange(2, headMap[boolCol] + 1).getDataValidation();
                    const existingBottom = sheet.getRange(lastRow - 1, headMap[boolCol] + 1).getDataValidation();
                    if (normalizeValidation_(existingTop) !== normalizeValidation_(rule) || normalizeValidation_(existingBottom) !== normalizeValidation_(rule)) {
                        plan.validationsToApply.push({ sheetName, col: headMap[boolCol] + 1, type: 'CHECKBOX' });
                    }

                 }
            }

            // Apply Prime / Special Week if found
            if (headMap['Prime Classification'] !== undefined) {
                 for (let r = 1; r < data.length; r++) {
                     const val = data[r][headMap['Prime Classification']];
                     if (val !== '' && val !== 'Prime' && val !== 'Non-Prime') {
                         plan.conflicts.push(`Existing invalid Prime Classification found at row ${r+1}`);
                     }
                 }
                 const rule = buildValidationRule_('DROPDOWN_PRIME');
                 const lastRow = Math.max(sheet.getMaxRows(), 2);
                 const existingTop = sheet.getRange(2, headMap['Prime Classification'] + 1).getDataValidation();
                 const existingBottom = sheet.getRange(lastRow - 1, headMap['Prime Classification'] + 1).getDataValidation();
                 if (normalizeValidation_(existingTop) !== normalizeValidation_(rule) || normalizeValidation_(existingBottom) !== normalizeValidation_(rule)) {
                     plan.validationsToApply.push({ sheetName, col: headMap['Prime Classification'] + 1, type: 'DROPDOWN_PRIME' });
                 }
            }
            if (headMap['Special Week'] !== undefined) {
                 for (let r = 1; r < data.length; r++) {
                     const val = data[r][headMap['Special Week']];
                     if (val !== '' && val !== 'None' && val !== 'Spring Break' && val !== 'Christmas') {
                         plan.conflicts.push(`Existing invalid Special Week found at row ${r+1}`);
                     }
                 }
                 const rule = buildValidationRule_('DROPDOWN_SPECIAL');
                 const lastRow = Math.max(sheet.getMaxRows(), 2);
                 const existingTop = sheet.getRange(2, headMap['Special Week'] + 1).getDataValidation();
                 const existingBottom = sheet.getRange(lastRow - 1, headMap['Special Week'] + 1).getDataValidation();
                 if (normalizeValidation_(existingTop) !== normalizeValidation_(rule) || normalizeValidation_(existingBottom) !== normalizeValidation_(rule)) {
                     plan.validationsToApply.push({ sheetName, col: headMap['Special Week'] + 1, type: 'DROPDOWN_SPECIAL' });
                 }
            }
        }
    }

    // Admin Options Default Defaults
    let adminOptionsSheet = ss.getSheetByName('Admin Options');
    if (adminOptionsSheet) {
        const adData = adminOptionsSheet.getDataRange().getValues();
        const adMap = _getHMap(adData);
        if (adMap['Setting'] !== undefined && adMap['Value'] !== undefined) {
             const existingSets = new Set();
             for (let i = 1; i < adData.length; i++) {
                  const s = String(adData[i][adMap['Setting']] || '').trim();
                  if (s) existingSets.add(s);
             }
             const defaults = {
                 'Default Vacation Week Target': '9',
                 'Default Vacation Week Capacity': '4',
                 'Vacation ACTIVE-window size': '3'
             };
             // Attach it to plan so we can append them
             plan.adminOptionsToAdd = {};
             for (const k in defaults) {
                 if (!existingSets.has(k)) {
                     plan.adminOptionsToAdd[k] = defaults[k];
                 }
             }

             // Validate 'Vacation ACTIVE-window size' is a STRICT_POS_INT
             const valCol = adMap['Value'];
             for (let i = 1; i < adData.length; i++) {
                 const key = String(adData[i][adMap['Setting']] || '').trim();
                 if (key === 'Vacation ACTIVE-window size' || key === 'Default Vacation Week Capacity' || key === 'Default Vacation Week Target') {
                     const val = adData[i][valCol];
                     if (val !== '' && (!Number.isInteger(Number(val)) || Number(val) <= 0)) {
                         plan.conflicts.push(`Existing invalid strictly positive integer found for Admin Option '${key}' at row ${i+1}`);
                     }
                 }
             }
        }
    }

    // Config
    let configSheet = ss.getSheetByName('Config');
    if (configSheet) {
        const cData = configSheet.getDataRange().getValues();
        const cMap = _getHMap(cData);
        if (cMap['Key'] !== undefined) {
            const existingKeys = new Set();
            const dups = new Set();
            for (let i = 1; i < cData.length; i++) {
                const k = String(cData[i][cMap['Key']] || '').trim();
                if (k) {
                    if (existingKeys.has(k)) dups.add(k);
                    existingKeys.add(k);
                }
            }
            if (dups.size > 0) {
                plan.conflicts.push(`Duplicate Config keys found: ${Array.from(dups).join(', ')}`);
            } else {
                for (const k in _DEFAULT_CONFIG_REF) {
                    if (!existingKeys.has(k)) {
                        plan.configKeysToAdd[k] = _DEFAULT_CONFIG_REF[k];

                    }
                }
            }
            if (cMap['Value'] !== undefined) {
                 for (let i = 1; i < cData.length; i++) {
                     const k = String(cData[i][cMap['Key']] || '').trim();
                     let ruleType = null;

                     if (k === 'Setup State') ruleType = 'DROPDOWN_SETUP';
                     else if (k === 'Current Phase') ruleType = 'DROPDOWN_PHASE';
                     else if (k === 'Phase Ready State') ruleType = 'DROPDOWN_READY';
                     else if (k === 'Current Serpentine Direction') ruleType = 'DROPDOWN_DIR';
                     else if (k === 'Active Year') {
                          ruleType = 'YEAR';
                          const val = cData[i][cMap['Value']];
                          if (val !== '' && !/^\d{4}$/.test(String(val))) plan.conflicts.push('Invalid Active Year format');
                     }
                     else if (['Current Vacation Round', 'Current Queue Cycle', 'Active Window Generation', 'Current Queue Cursor'].includes(k)) {
                          ruleType = 'POS_INT';
                          const val = cData[i][cMap['Value']];
                          if (val !== '' && (!Number.isInteger(Number(val)) || Number(val) < 0)) plan.conflicts.push(`Invalid config number for ${k}`);
                     }

                     if (ruleType) {
                         const rule = buildValidationRule_(ruleType);
                         const existingConf = configSheet.getRange(i+1, cMap['Value']+1).getDataValidation();
                         if (normalizeValidation_(existingConf) !== normalizeValidation_(rule)) {
                             plan.validationsToApply.push({ sheetName: 'Config', row: i+1, col: cMap['Value']+1, type: ruleType });
                         }
                     }
                 }
                 // Handle newly added ones
                 let nextRow = cData.length + 1;
                 for (const k in plan.configKeysToAdd) {
                     if (k === 'Setup State') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: cMap['Value']+1, type: 'DROPDOWN_SETUP' });
                     if (k === 'Current Phase') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: cMap['Value']+1, type: 'DROPDOWN_PHASE' });
                     if (k === 'Phase Ready State') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: cMap['Value']+1, type: 'DROPDOWN_READY' });
                     if (k === 'Current Serpentine Direction') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: cMap['Value']+1, type: 'DROPDOWN_DIR' });
                     if (k === 'Active Year') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: cMap['Value']+1, type: 'YEAR' });
                     if (['Current Vacation Round', 'Current Queue Cycle', 'Active Window Generation', 'Current Queue Cursor'].includes(k)) {
                         plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: cMap['Value']+1, type: 'POS_INT' });
                     }
                     nextRow++;
                 }
            }
        }
    } else {
        // If config sheet missing, keys are added, validations need to be added too

        for (const k in _DEFAULT_CONFIG_REF) { plan.configKeysToAdd[k] = _DEFAULT_CONFIG_REF[k]; }
        let nextRow = 2; // Row 1 is header
        for (const k in _DEFAULT_CONFIG_REF) {
            if (k === 'Setup State') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: 2, type: 'DROPDOWN_SETUP' });
            if (k === 'Current Phase') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: 2, type: 'DROPDOWN_PHASE' });
            if (k === 'Phase Ready State') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: 2, type: 'DROPDOWN_READY' });
            if (k === 'Current Serpentine Direction') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: 2, type: 'DROPDOWN_DIR' });
            if (k === 'Active Year') plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: 2, type: 'YEAR' });
            if (['Current Vacation Round', 'Current Queue Cycle', 'Active Window Generation', 'Current Queue Cursor'].includes(k)) {
                 plan.validationsToApply.push({ sheetName: 'Config', row: nextRow, col: 2, type: 'POS_INT' });
            }
            nextRow++;
        }
    }


    plan.hasChanges = plan.sheetsToAdd.length > 0 ||
                      Object.keys(plan.headersToAppend).length > 0 ||
                      Object.keys(plan.configKeysToAdd).length > 0 ||
                      plan.validationsToApply.length > 0 ||
                      plan.formatsToApply.length > 0;
    return plan;

}


function normalizeValidation_(rule) {
    if (!rule) return null;
    // Mock might just be a plain object, try stringifying if it doesn't have AS methods
    if (typeof rule.getCriteriaType !== 'function') {
        // Assume it's a mock plain object builder
        return JSON.stringify({
            criteriaType: rule.criteriaType,
            args: rule.args,
            helpText: rule.helpText,
            allowInvalid: rule._allowInvalid
        });
    }
    // Prod AS
    return JSON.stringify({
        criteriaType: String(rule.getCriteriaType()),
        args: rule.getCriteriaValues(),
        helpText: rule.getHelpText(),
        allowInvalid: rule.getAllowInvalid()
    });
}

function buildValidationRule_(type) {
    if (type === 'CHECKBOX') return SpreadsheetApp.newDataValidation().requireCheckbox().build();
    if (type === 'POS_INT') return SpreadsheetApp.newDataValidation().requireFormulaSatisfied(`=OR(ISBLANK(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)), AND(ISNUMBER(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)), INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)>=0, INT(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE))=INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)))`).setHelpText('Must be a blank or a positive whole number.').build();
    if (type === 'STRICT_POS_INT') return SpreadsheetApp.newDataValidation().requireFormulaSatisfied(`=OR(ISBLANK(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)), AND(ISNUMBER(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)), INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)>0, INT(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE))=INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)))`).setHelpText('Must be a blank or a positive whole number greater than 0.').build();
    if (type === 'YEAR') return SpreadsheetApp.newDataValidation().requireFormulaSatisfied(`=OR(ISBLANK(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)), AND(ISNUMBER(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE)), LEN(INDIRECT("R"&ROW()&"C"&COLUMN(), FALSE))=4))`).setHelpText('Must be a blank or 4-digit year.').build();
    if (type === 'LIST_PRIME') return SpreadsheetApp.newDataValidation().requireValueInList(['Prime', 'Non-Prime']).build();
    if (type === 'LIST_SPECIAL') return SpreadsheetApp.newDataValidation().requireValueInList(['None', 'Spring Break', 'Christmas']).build();

    if (type === 'DROPDOWN_SETUP') return SpreadsheetApp.newDataValidation().requireValueInList(_VALID_SETUP_STATES).build();
    if (type === 'DROPDOWN_PHASE') return SpreadsheetApp.newDataValidation().requireValueInList(_VALID_PHASES).build();
    if (type === 'DROPDOWN_READY') return SpreadsheetApp.newDataValidation().requireValueInList(_VALID_READY_STATES).build();
    if (type === 'DROPDOWN_DIR') return SpreadsheetApp.newDataValidation().requireValueInList(_VALID_SERPENTINE_DIRECTIONS).build();

    return null;
}

function initializeOrUpdateWorkbook_() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return _apiResponse(false, null, 'System is busy.');
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const plan = planSchema_(ss);

    if (plan.conflicts.length > 0) {
        return _apiResponse(false, { report: plan, changed: false }, 'Initialization aborted due to conflicts.');
    }

    // Always apply validations and formats as we cannot easily read existing ones to check if changed
    // Apply changes
    for (const sheetName of plan.sheetsToAdd) {
        ss.insertSheet(sheetName);
    }

    for (const sheetName in plan.headersToAppend) {
        const sheet = ss.getSheetByName(sheetName);
        const missing = plan.headersToAppend[sheetName];
        if (missing.length > 0) {
            const numCols = sheet.getLastColumn();
            if (numCols === 0) {
                sheet.getRange(1, 1, 1, missing.length).setValues([missing]);
            } else {
                sheet.getRange(1, numCols + 1, 1, missing.length).setValues([missing]);
            }
        }
    }

    if (plan.adminOptionsToAdd && Object.keys(plan.adminOptionsToAdd).length > 0) {
        const adminSh = ss.getSheetByName('Admin Options');
        const adData = adminSh.getDataRange().getValues();
        const adMap = _getHMap(adData);
        const keysToAdd = Object.keys(plan.adminOptionsToAdd);
        let lastRow = adminSh.getLastRow();
        const rowsToAppend = keysToAdd.map(k => {
             const row = new Array(Math.max(adminSh.getLastColumn(), 2)).fill('');
             row[adMap['Setting']] = k;
             row[adMap['Value']] = plan.adminOptionsToAdd[k];
             return row;
        });
        adminSh.getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
    }

    const configSh = ss.getSheetByName('Config');
    if (configSh) {
        const cData = configSh.getDataRange().getValues();
        const cMap = _getHMap(cData);
        if (cMap['Key'] !== undefined && cMap['Value'] !== undefined) {
            const keysToAdd = Object.keys(plan.configKeysToAdd);
            if (keysToAdd.length > 0) {
                let lastRow = configSh.getLastRow();
                const rowsToAppend = keysToAdd.map(k => {
                   const row = new Array(Math.max(configSh.getLastColumn(), 2)).fill('');
                   row[cMap['Key']] = k;
                   row[cMap['Value']] = plan.configKeysToAdd[k];
                   return row;
                });
                configSh.getRange(lastRow + 1, 1, rowsToAppend.length, rowsToAppend[0].length).setValues(rowsToAppend);
            }
        }
    }

    // Apply formats and validations
    for (const v of plan.formatsToApply) {
        const sheet = ss.getSheetByName(v.sheetName);
        if (sheet) {
             const lastRow = Math.max(sheet.getMaxRows(), 2);
             sheet.getRange(2, v.col, lastRow - 1, 1).setNumberFormat(v.format);
        }
    }

    for (const v of plan.validationsToApply) {
        const sheet = ss.getSheetByName(v.sheetName);
        if (sheet) {
            const rule = buildValidationRule_(v.type);
            if (rule) {
                if (v.row) {
                    sheet.getRange(v.row, v.col).setDataValidation(rule);
                } else {
                    const lastRow = Math.max(sheet.getMaxRows(), 2);
                    sheet.getRange(2, v.col, lastRow - 1, 1).setDataValidation(rule);
                }
            }
        }
    }

    // Check if we did anything
    // Validations and formats might be reapplied, but we consider "changed" if structural things or keys were added.

    if (!plan.hasChanges) {
        return _apiResponse(true, { report: plan, changed: false }, 'Workbook is already fully initialized.');
    }
    return _apiResponse(true, { report: plan, changed: true }, 'Workbook schema initialized and updated successfully.');


  } catch (e) {
      return _apiResponse(false, null, 'Unexpected error during initialization: ' + e.message);
  } finally {
      lock.releaseLock();
  }
}

function inspectSchema_(token) {
    const _reqAdmin = typeof requireAdmin_ === 'function' ? requireAdmin_ : (typeof global !== 'undefined' && global.requireAdmin_ ? global.requireAdmin_ : (require('./Auth.gs').requireAdmin));
    if (!_reqAdmin(token)) return _apiResponse(false, null, 'Unauthorized');

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
        return _apiResponse(false, null, 'System is busy processing another request.');
    }

    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        return _apiResponse(true, planSchema_(ss), 'Schema inspected');
    } catch(e) {
        return _apiResponse(false, null, 'Unexpected error during inspection: ' + e.message);
    } finally {
        lock.releaseLock();
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        planSchema: planSchema_,
        inspectSchema: inspectSchema_,
        initializeOrUpdateWorkbook: initializeOrUpdateWorkbook_,
        REQUIRED_SHEETS,
        REQUIRED_HEADERS
    };
}
