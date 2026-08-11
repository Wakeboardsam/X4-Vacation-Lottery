// Dependencies
const _vState = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));
const _vApi = typeof apiResponse_ === 'function' ? apiResponse_ : (typeof global !== 'undefined' && global.apiResponse_ ? global.apiResponse_ : (require('./Utils.gs').apiResponse));

function _vWrite(updates) {
    const writer = typeof writeConfigState_ === 'function' ? writeConfigState_ : (typeof global !== 'undefined' && global.writeConfigState_ ? global.writeConfigState_ : (require('./State.gs').writeConfigState));
    return writer(updates);
}
const _vReqAuth = typeof resolveParticipantSession_ === 'function' ? resolveParticipantSession_ : (typeof global !== 'undefined' && global.resolveParticipantSession_ ? global.resolveParticipantSession_ : (require('./Auth.gs').resolveParticipantSession));
const _vReqAdmin = typeof requireAdmin_ === 'function' ? requireAdmin_ : (typeof global !== 'undefined' && global.requireAdmin_ ? global.requireAdmin_ : (require('./Auth.gs').requireAdmin));
const _vHMap = typeof getHeaderMap_ === 'function' ? getHeaderMap_ : (typeof global !== 'undefined' && global.getHeaderMap_ ? global.getHeaderMap_ : (require('./Utils.gs').getHeaderMap));
const _vFindRow = typeof findRowIndex_ === 'function' ? findRowIndex_ : (typeof global !== 'undefined' && global.findRowIndex_ ? global.findRowIndex_ : (require('./Utils.gs').findRowIndex));
const _vRequireHeaders = typeof requireHeaders_ === 'function' ? requireHeaders_ : (typeof global !== 'undefined' && global.requireHeaders_ ? global.requireHeaders_ : (require('./Utils.gs').requireHeaders));
const _vValidateParticipantIds = typeof validateParticipantIds_ === 'function' ? validateParticipantIds_ : (typeof global !== 'undefined' && global.validateParticipantIds_ ? global.validateParticipantIds_ : (require('./Utils.gs').validateParticipantIds));
const _vCalcNext = typeof calculateNextQueueState_ === 'function' ? calculateNextQueueState_ : (typeof global !== 'undefined' && global.calculateNextQueueState_ ? global.calculateNextQueueState_ : (require('./QueueEngine.gs').calculateNextQueueState));

function parseRequiredPositiveWholeNumber_(value, label) {
    const raw = String(value === null || value === undefined ? '' : value).trim();
    const parsed = Number(raw);
    if (raw === '' || !Number.isInteger(parsed) || parsed <= 0) {
        const err = new Error(`${label} must be a positive whole number. Correct Admin Options before continuing.`);
        err.code = 'ADMIN_CONFIGURATION_ERROR';
        throw err;
    }
    return parsed;
}

function normalizeWeekStartDate_(value) {
    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
        if (typeof Utilities !== 'undefined' && typeof Utilities.formatDate === 'function') {
            let timeZone = 'Etc/UTC';
            const ss = SpreadsheetApp.getActiveSpreadsheet();
            if (ss && typeof ss.getSpreadsheetTimeZone === 'function') timeZone = ss.getSpreadsheetTimeZone();
            return Utilities.formatDate(value, timeZone, 'yyyy-MM-dd');
        }
        const year = value.getFullYear();
        const month = String(value.getMonth() + 1).padStart(2, '0');
        const day = String(value.getDate()).padStart(2, '0');
        return `${year}-${month}-${day}`;
    }
    return String(value === null || value === undefined ? '' : value).trim();
}

function getAdminOptions_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Admin Options');
    if (!sheet) throw new Error('Admin Options sheet is missing. Initialize the workbook before continuing.');

    const data = sheet.getDataRange().getValues();
    const map = _vHMap(data);
    _vRequireHeaders(map, ['Setting', 'Value'], 'Admin Options');
    const settings = {};

    if (map['Setting'] !== undefined && map['Value'] !== undefined) {
        for (let i = 1; i < data.length; i++) {
            const key = String(data[i][map['Setting']]).trim();
            if (key) settings[key] = data[i][map['Value']];
        }
    }
    const target = parseRequiredPositiveWholeNumber_(settings['Default Vacation Week Target'], 'Default Vacation Week Target');
    const capacity = parseRequiredPositiveWholeNumber_(settings['Default Vacation Week Capacity'], 'Default Vacation Week Capacity');
    const windowSize = parseRequiredPositiveWholeNumber_(settings['Vacation ACTIVE-window size'], 'Vacation ACTIVE-window size');
    return { target, capacity, windowSize };
}

function getRosterForVacation_(activeYear, globalTarget) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Turn Management');
    if (!sheet) throw new Error('Turn Management sheet is missing.');
    const data = sheet.getDataRange().getValues();
    const map = _vHMap(data);
    _vRequireHeaders(map, [
        'Name', 'Participant ID', 'PIN', 'Active for Year', 'Vacation Phase Enabled',
        'Vacation Week Target Override', 'Seniority Position', 'Lottery Position',
        'Had Spring Break Last Year', 'Had Christmas Week Last Year'
    ], 'Turn Management');
    _vValidateParticipantIds(data, map);

    const roster = [];
    for (let i = 1; i < data.length; i++) {
        const pId = String(data[i][map['Participant ID']] || '').trim();
        const name = String(data[i][map['Name']] || '').trim();
        const pin = String(data[i][map['PIN']] || '').trim();
        if (!name && !pId && !pin) continue;

        const active = String(data[i][map['Active for Year']] || '').toLowerCase() === 'true';
        const phaseEnabled = String(data[i][map['Vacation Phase Enabled']] || '').toLowerCase() === 'true';

        let disposition = 'EXCLUDED';
        if (active && phaseEnabled) {
             disposition = 'ELIGIBLE';
        }

        let override = String(data[i][map['Vacation Week Target Override']] || '').trim();
        const target = override === '' ? globalTarget : parseRequiredPositiveWholeNumber_(override, `Vacation Week Target Override for ${name || pId}`);

        roster.push({
            participantId: pId,
            name: name,
            seniority: Number(String(data[i][map['Seniority Position']] || '').trim()),
            lottery: Number(String(data[i][map['Lottery Position']] || '').trim()),
            target: target,
            hadSpringBreak: String(data[i][map['Had Spring Break Last Year']] || '').toLowerCase() === 'true',
            hadChristmas: String(data[i][map['Had Christmas Week Last Year']] || '').toLowerCase() === 'true',
            disposition: disposition,
            rowIndex: i
        });
    }
    return roster;
}

function getWeekAvailability_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Week Availability');
    if (!sheet) return { data: [], map: {}, sheet };

    const data = sheet.getDataRange().getValues();
    const map = _vHMap(data);
    _vRequireHeaders(map, ['Week Start Date', 'Prime Classification', 'Special Week', 'Capacity Override'], 'Week Availability');
    return { data, map, sheet };
}


function submitVacation_(participantId, submittedTurnId, selections) {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return _vApi(false, null, 'System is busy processing another request.');

    try {
        const config = _vState();
        if (config['Current Phase'] !== 'VACATION_SENIORITY' && config['Current Phase'] !== 'VACATION_RANDOM') {
            return _vApi(false, null, 'Vacation phase is not active.');
        }
        if (config['Phase Ready State'] === 'READY_WEEKEND') {
            return _vApi(false, null, 'Vacation selection is closed. Weekend is ready for the administrator to start.');
        }

        const activeWindow = JSON.parse(config['Current Active Window'] || '[]');
        const targetTurn = activeWindow.find(t => t.turnId === submittedTurnId);
        if (!targetTurn) return _vApi(false, null, 'Your turn is no longer active (it may have timed out or you are using a stale screen).');
        if (targetTurn.participantId !== participantId) return _vApi(false, null, 'Turn mismatch.');

        const adminOpts = getAdminOptions_();
        const globalTarget = adminOpts.target;
        const globalCapacity = adminOpts.capacity;

        const roster = getRosterForVacation_(config['Active Year'], globalTarget);
        const pRoster = roster.find(r => r.participantId === participantId);
        if (!pRoster || pRoster.disposition !== 'ELIGIBLE') {
            return _vApi(false, null, 'You are not eligible to select vacation.');
        }

        const { data: weekData, map: weekMap, sheet: weekSheet } = getWeekAvailability_();

        // Count participant's existing selections
        let currentSelections = 0;
        let existingWeeks = new Set();
        let maxPersonColFound = 0;

        const personCols = Object.keys(weekMap).filter(k => k.startsWith('Person')).sort();
        if (personCols.length > 0) {
            maxPersonColFound = Number(personCols[personCols.length - 1].replace('Person', ''));
        }

        for (let r = 1; r < weekData.length; r++) {
            for (const col of personCols) {
                if (String(weekData[r][weekMap[col]] || '').trim() === pRoster.name) {
                     currentSelections++;
                     existingWeeks.add(normalizeWeekStartDate_(weekData[r][weekMap['Week Start Date']]));
                }
            }
        }

        if (currentSelections >= pRoster.target) {
            return _vApi(false, null, 'You have already reached your vacation target.');
        }

        if (!Array.isArray(selections) || selections.length === 0 || selections.length > 2) {
             return _vApi(false, null, 'You must select exactly one or two weeks.');
        }

        const normalizedSelections = selections.map(normalizeWeekStartDate_);
        if (normalizedSelections.some(value => value === '')) {
            return _vApi(false, null, 'Every selected vacation week must have a valid Week Start Date.');
        }

        if (currentSelections + normalizedSelections.length > pRoster.target) {
             return _vApi(false, null, 'Selections exceed your vacation target.');
        }

        const roundNum = Number(config['Current Vacation Round'] || 0);
        let primeCount = 0;
        let selectedRowUpdates = [];
        let requiredMaxPersonCol = maxPersonColFound;

        // Validate each selection
        for (const selWeek of normalizedSelections) {
             if (existingWeeks.has(selWeek)) {
                 return _vApi(false, null, `You already hold week ${selWeek}.`);
             }
             if (normalizedSelections.filter(s => s === selWeek).length > 1) {
                 return _vApi(false, null, `Duplicate selection for week ${selWeek}.`);
             }

             let wRowIndex = -1;
             for (let rowIndex = 1; rowIndex < weekData.length; rowIndex++) {
                 if (normalizeWeekStartDate_(weekData[rowIndex][weekMap['Week Start Date']]) === selWeek) {
                     wRowIndex = rowIndex;
                     break;
                 }
             }
             if (wRowIndex === -1) {
                 return _vApi(false, null, `Week ${selWeek} not found.`);
             }

             const row = weekData[wRowIndex];
             const primeType = String(row[weekMap['Prime Classification']] || '').trim();
             const specialType = String(row[weekMap['Special Week']] || '').trim();
             const capOverride = String(row[weekMap['Capacity Override']] || '').trim();
             const capacity = capOverride === '' ? globalCapacity : parseRequiredPositiveWholeNumber_(capOverride, `Capacity Override for week ${selWeek}`);

             if (primeType === 'Prime') primeCount++;

             if (roundNum <= 3) {
                 if (specialType === 'Spring Break' && pRoster.hadSpringBreak) {
                     return _vApi(false, null, 'Cannot select Spring Break in rounds 1-3 due to prior year rules.');
                 }
                 if (specialType === 'Christmas' && pRoster.hadChristmas) {
                     return _vApi(false, null, 'Cannot select Christmas in rounds 1-3 due to prior year rules.');
                 }
             }

             let assignedCount = 0;
             let emptyPersonCol = -1;
             for (let c = 1; c <= Math.max(capacity, maxPersonColFound); c++) {
                 const colName = `Person${c}`;
                 if (weekMap[colName] !== undefined && String(row[weekMap[colName]] || '').trim() !== '') {
                     assignedCount++;
                 } else if (emptyPersonCol === -1) {
                     emptyPersonCol = c;
                 }
             }

             if (assignedCount >= capacity) {
                 return _vApi(false, null, `Week ${selWeek} is at maximum capacity.`);
             }

             if (emptyPersonCol > requiredMaxPersonCol) requiredMaxPersonCol = emptyPersonCol;

             selectedRowUpdates.push({
                 rowIndex: wRowIndex,
                 colToUpdate: emptyPersonCol
             });
        }

        if (primeCount > 1) {
             return _vApi(false, null, 'You may only select a maximum of one Prime week per turn.');
        }
        if (primeCount === 1 && normalizedSelections.length > 1) {
             return _vApi(false, null, 'A Prime week must stand alone (no additional Non-Prime picks).');
        }

        // Capture original week data for potential rollback
        const originalWeekData = weekData.map(row => row.slice());

        // Passed all validation. Need to expand capacity columns if needed
        if (requiredMaxPersonCol > maxPersonColFound) {
            const missing = [];
            for (let c = maxPersonColFound + 1; c <= requiredMaxPersonCol; c++) {
                 missing.push(`Person${c}`);
            }
            if (missing.length > 0) {
                const numCols = weekSheet.getLastColumn();
                weekSheet.getRange(1, numCols + 1, 1, missing.length).setValues([missing]);

                // Keep memory matrix aligned by appending the new headers and padding the rows
                for (let i = 0; i < missing.length; i++) {
                     weekMap[missing[i]] = numCols + i;
                     originalWeekData[0].push(missing[i]); // Keep original matrix aligned in shape
                }

                for (let i = 0; i < weekData.length; i++) {
                    for (let c = 0; c < missing.length; c++) {
                         weekData[i].push("");
                         if (i > 0) originalWeekData[i].push("");
                    }
                }
                weekData[0] = Object.keys(weekMap).sort((a,b) => weekMap[a] - weekMap[b]);
            }
        }

        // Make assignments in memory
        for (const upd of selectedRowUpdates) {
             const colName = `Person${upd.colToUpdate}`;
             weekData[upd.rowIndex][weekMap[colName]] = pRoster.name;
        }

        // Determine Next Queue State
        const twoNonPrime = primeCount === 0 && normalizedSelections.length === 2;
        let newSelectionsTotal = currentSelections + normalizedSelections.length;
        if (newSelectionsTotal >= pRoster.target) {
             pRoster.disposition = 'COMPLETE_FOR_PHASE';
        }

        const queueConfig = {
            movementMode: config['Current Phase'] === 'VACATION_SENIORITY' ? 'FORWARD_ONLY' : 'SERPENTINE',
            orderSource: config['Current Phase'] === 'VACATION_SENIORITY' ? 'seniority' : 'lottery',
            windowSize: adminOpts.windowSize,
            action: 'COMPLETE',
            completedTurnId: submittedTurnId,
            phase: config['Current Phase']
        };

        // Refresh roster dispositions based on new totals
        for (let i = 0; i < roster.length; i++) {
            if (roster[i].participantId === participantId) {
                roster[i].disposition = pRoster.disposition;
            } else if (roster[i].disposition === 'ELIGIBLE') {
                // Determine if they also reached target previously
                let otherSelections = 0;
                for (let r = 1; r < weekData.length; r++) {
                    for (const col of Object.keys(weekMap).filter(k => k.startsWith('Person'))) {
                        if (String(weekData[r][weekMap[col]] || '').trim() === roster[i].name) {
                            otherSelections++;
                        }
                    }
                }
                if (otherSelections >= roster[i].target) {
                    roster[i].disposition = 'COMPLETE_FOR_PHASE';
                }
            }
        }

        // Prepare parsed state for queue engine
        const parsedState = {
            ...config,
            "Current Active Window": JSON.parse(config["Current Active Window"] || "[]"),
            "Current Directional Window": JSON.parse(config["Current Directional Window"] || "[]"),
            "Current Directional Window Completed": JSON.parse(config["Current Directional Window Completed"] || "[]"),
            "Current Queue Skip State": JSON.parse(config["Current Queue Skip State"] || "{}"),
            "Current Queue Cursor": Number(config["Current Queue Cursor"] || 0),
            "Current Queue Cycle": Number(config["Current Queue Cycle"] || 0),
            "Active Window Generation": Number(config["Active Window Generation"] || 0)
        };

        if (twoNonPrime) {
            parsedState["Current Queue Skip State"][participantId] = (parsedState["Current Queue Skip State"][participantId] || 0) + 1;
        }

        const { nextState, queueComplete } = _vCalcNext(parsedState, roster, queueConfig);

        // Phase transition checks
        if (queueComplete) {
            if (config['Current Phase'] === 'VACATION_SENIORITY') {
                 // Move to Round 2
                 nextState['Current Phase'] = 'VACATION_RANDOM';
                 nextState['Current Vacation Round'] = '2';
                 nextState['Current Serpentine Direction'] = 'FORWARD';

                 // Clean restart config
                 nextState["Current Active Window"] = [];
                 nextState["Current Directional Window"] = [];
                 nextState["Current Directional Window Completed"] = [];
                 nextState["Current Queue Cursor"] = 0;
                 nextState["Current Queue Cycle"] = 0;

                 const p2Config = {
                    movementMode: 'SERPENTINE',
                    orderSource: 'lottery',
                    windowSize: adminOpts.windowSize,
                    action: 'INIT',
                    phase: 'VACATION_RANDOM',
                    preserveSkips: true
                 };
                 // Set all dispositions strictly to ELIGIBLE to re-initialize clean unless they completed phase
                 for (let r of roster) {
                    if (r.disposition !== 'COMPLETE_FOR_PHASE') {
                        r.disposition = 'ELIGIBLE';
                    }
                 }
                 const p2State = _vCalcNext(nextState, roster, p2Config);
                 Object.assign(nextState, p2State.nextState);
                 nextState['Current Vacation Round'] = String(1 + Number(nextState['Current Queue Cycle'] || 1));
                 if (p2State.queueComplete) {
                     nextState['Phase Ready State'] = 'READY_WEEKEND';
                     nextState["Current Active Window"] = [];
                     nextState["Current Directional Window"] = [];
                     nextState["Current Directional Window Completed"] = [];
                 }

            } else {
                 // Fully complete
                 nextState['Phase Ready State'] = 'READY_WEEKEND';
                 nextState["Current Active Window"] = [];
                 nextState["Current Directional Window"] = [];
                 nextState["Current Directional Window Completed"] = [];
            }
        } else if (config['Current Phase'] === 'VACATION_RANDOM') {
            // Check if cycle advanced to increment round
            if (nextState["Current Queue Cycle"] > parsedState["Current Queue Cycle"]) {
                 nextState['Current Vacation Round'] = (Number(nextState['Current Vacation Round']) + 1).toString();
            }
        }

        const updates = {
            "Current Phase": nextState["Current Phase"],
            "Phase Ready State": nextState["Phase Ready State"],
            "Current Vacation Round": nextState["Current Vacation Round"],
            "Current Queue Phase": nextState["Current Queue Phase"],
            "Current Queue Order Source": nextState["Current Queue Order Source"],
            "Current Active Window": JSON.stringify(nextState["Current Active Window"]),
            "Current Directional Window": JSON.stringify(nextState["Current Directional Window"]),
            "Current Directional Window Completed": JSON.stringify(nextState["Current Directional Window Completed"]),
            "Current Queue Skip State": JSON.stringify(nextState["Current Queue Skip State"] || {}),
            "Current Queue Cursor": nextState["Current Queue Cursor"].toString(),
            "Current Queue Cycle": nextState["Current Queue Cycle"].toString(),
            "Current Serpentine Direction": nextState["Current Serpentine Direction"],
            "Active Window Generation": nextState["Active Window Generation"].toString()
        };

        // Atomic writes:
        // Write week data matrix all at once
        weekSheet.getRange(1, 1, weekData.length, weekData[0].length).setValues(weekData);

        // Then write config
        try {
            _vWrite(updates);
        } catch(writeErr) {
            try {
                weekSheet.getRange(1, 1, originalWeekData.length, originalWeekData[0].length).setValues(originalWeekData);
                if (typeof Logger !== 'undefined' && Logger.log) {
                    Logger.log('Vacation submission failed while saving queue state. Week Availability rollback succeeded; no changes were committed. ' + writeErr.message);
                }
                return _vApi(false, null, 'Your vacation selection could not be saved. No changes were made. Please try again.');
            } catch(rollbackErr) {
                const fatal = new Error('Fatal consistency error: queue state save failed and Week Availability rollback also failed. Stop vacation processing and inspect the workbook before continuing. Config error: ' + writeErr.message + '; rollback error: ' + rollbackErr.message);
                fatal.code = 'FATAL_CONSISTENCY_ERROR';
                throw fatal;
            }
        }

        return _vApi(true, {
             targetReached: newSelectionsTotal >= pRoster.target,
             queueComplete: queueComplete && config['Current Phase'] !== 'VACATION_SENIORITY'
        }, 'Vacation weeks successfully selected.');

    } catch(e) {
        if (e.code === 'FATAL_CONSISTENCY_ERROR') throw e;
        if (e.code === 'PARTICIPANT_ID_CONFLICT') return _vApi(false, null, e.participantMessage);
        if (e.code === 'ADMIN_CONFIGURATION_ERROR') return _vApi(false, null, 'Vacation selection is temporarily unavailable because an administrator setting requires correction. No changes were made.');
        return _vApi(false, null, 'Error: ' + e.message);
    } finally {
        lock.releaseLock();
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        submitVacation: submitVacation_,
        getRosterForVacation: getRosterForVacation_,
        getAdminOptions: getAdminOptions_,
        getWeekAvailability: getWeekAvailability_,
        normalizeWeekStartDate: normalizeWeekStartDate_
    };
}
