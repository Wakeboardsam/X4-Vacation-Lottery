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
const _vCalcNext = typeof calculateNextQueueState_ === 'function' ? calculateNextQueueState_ : (typeof global !== 'undefined' && global.calculateNextQueueState_ ? global.calculateNextQueueState_ : (require('./QueueEngine.gs').calculateNextQueueState));

function getAdminOptions_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Admin Options');
    if (!sheet) return { target: 9, capacity: 4, windowSize: 3 };

    const data = sheet.getDataRange().getValues();
    const map = _vHMap(data);
    let target = 9;
    let capacity = 4;
    let windowSize = 3;

    if (map['Setting'] !== undefined && map['Value'] !== undefined) {
        for (let i = 1; i < data.length; i++) {
            const key = String(data[i][map['Setting']]).trim();
            if (key === 'Default Vacation Week Target') {
                target = Number(data[i][map['Value']]) || 9;
            } else if (key === 'Default Vacation Week Capacity') {
                capacity = Number(data[i][map['Value']]) || 4;
            } else if (key === 'Vacation ACTIVE-window size') {
                windowSize = Number(data[i][map['Value']]) || 3;
            }
        }
    }
    return { target, capacity, windowSize };
}

function getRosterForVacation_(activeYear, globalTarget) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Turn Management');
    const data = sheet.getDataRange().getValues();
    const map = _vHMap(data);

    const roster = [];
    for (let i = 1; i < data.length; i++) {
        const pId = String(data[i][map['PIN']] || '').trim();
        if (!pId) continue;

        const active = String(data[i][map['Active for Year']] || '').toLowerCase() === 'true';
        const phaseEnabled = String(data[i][map['Vacation Phase Enabled']] || '').toLowerCase() === 'true';

        let disposition = 'EXCLUDED';
        if (active && phaseEnabled) {
             disposition = 'ELIGIBLE';
        }

        let override = String(data[i][map['Vacation Week Target Override']] || '').trim();
        const target = override === '' ? globalTarget : Number(override);

        roster.push({
            participantId: pId,
            name: data[i][map['Name']],
            seniority: Number(data[i][map['Seniority Position']] || 999),
            lottery: Number(data[i][map['Lottery Position']] || 999),
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
                     existingWeeks.add(String(weekData[r][weekMap['Week Start Date']] || '').trim());
                }
            }
        }

        if (currentSelections >= pRoster.target) {
            return _vApi(false, null, 'You have already reached your vacation target.');
        }

        if (!Array.isArray(selections) || selections.length === 0 || selections.length > 2) {
             return _vApi(false, null, 'You must select exactly one or two weeks.');
        }

        if (currentSelections + selections.length > pRoster.target) {
             return _vApi(false, null, 'Selections exceed your vacation target.');
        }

        const roundNum = Number(config['Current Vacation Round'] || 0);
        let primeCount = 0;
        let selectedRowUpdates = [];
        let requiredMaxPersonCol = maxPersonColFound;

        // Validate each selection
        for (const selWeek of selections) {
             if (existingWeeks.has(selWeek)) {
                 return _vApi(false, null, `You already hold week ${selWeek}.`);
             }
             if (selections.filter(s => s === selWeek).length > 1) {
                 return _vApi(false, null, `Duplicate selection for week ${selWeek}.`);
             }

             const wRowIndex = _vFindRow(weekData, weekMap['Week Start Date'], selWeek);
             if (wRowIndex === -1) {
                 return _vApi(false, null, `Week ${selWeek} not found.`);
             }

             const row = weekData[wRowIndex];
             const primeType = String(row[weekMap['Prime Classification']] || '').trim();
             const specialType = String(row[weekMap['Special Week']] || '').trim();
             const capOverride = String(row[weekMap['Capacity Override']] || '').trim();
             const capacity = capOverride === '' ? globalCapacity : Number(capOverride);

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
        if (primeCount === 1 && selections.length > 1) {
             return _vApi(false, null, 'A Prime week must stand alone (no additional Non-Prime picks).');
        }

        // Capture original week data for potential rollback
        const originalWeekData = JSON.parse(JSON.stringify(weekData));

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
        const twoNonPrime = primeCount === 0 && selections.length === 2;
        let newSelectionsTotal = currentSelections + selections.length;
        if (newSelectionsTotal >= pRoster.target) {
             pRoster.disposition = 'COMPLETE_FOR_PHASE';
        } else if (twoNonPrime) {
             pRoster.disposition = 'SKIP_ONCE';
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

        // Fix for skips: if we just generated a skip, ensure it isn't consumed by the current window refill step
        // QueueEngine logic doesn't differentiate between old skips and newly generated ones for the currently processing turn.
        // It shouldn't consume the skip until the NEXT turn.
        if (twoNonPrime) {
             nextState["Current Queue Skip State"][participantId] = parsedState["Current Queue Skip State"][participantId];
        }

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

                 // Preserve skips for next phase
                 let nextSkips = { ...nextState["Current Queue Skip State"] };

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

                 // Restore skips - QueueEngine INIT wipes them intentionally if we don't carry them forward properly
                 nextState["Current Queue Skip State"] = nextSkips;

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
            // Rollback week assignments if config write fails
            weekSheet.getRange(1, 1, originalWeekData.length, originalWeekData[0].length).setValues(originalWeekData);
            throw new Error('Config save failed. Rollback applied. ' + writeErr.message);
        }

        return _vApi(true, {
             targetReached: newSelectionsTotal >= pRoster.target,
             queueComplete: queueComplete && config['Current Phase'] !== 'VACATION_SENIORITY'
        }, 'Vacation weeks successfully selected.');

    } catch(e) {
        if (e.message.includes('Rollback applied')) {
            throw e; // Fatal consistency errors should blow up as requested
        }
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
        getWeekAvailability: getWeekAvailability_
    };
}
