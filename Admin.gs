// Injection wrappers
const _apiR = typeof apiResponse_ === 'function' ? apiResponse_ : (typeof global !== 'undefined' && global.apiResponse_ ? global.apiResponse_ : (require('./Utils.gs').apiResponse));
const _readC = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));
const _initW = typeof initializeOrUpdateWorkbook_ === 'function' ? initializeOrUpdateWorkbook_ : (typeof global !== 'undefined' && global.initializeOrUpdateWorkbook_ ? global.initializeOrUpdateWorkbook_ : (require('./Schema.gs').initializeOrUpdateWorkbook));

/**
 * Validates the admin token before proceeding with any action
 * @param {string} token
 * @returns {boolean}
 */
function requireAdmin_(token) {
    const _resAdmin = typeof resolveAdminSession_ === 'function' ? resolveAdminSession_ : (typeof global !== 'undefined' && global.resolveAdminSession_ ? global.resolveAdminSession_ : (require('./Auth.gs').resolveAdminSession));
    return _resAdmin(token);
}

/**
 * Returns safe durable-state values for the dashboard
 * @param {string} token
 * @returns {object} API response
 */
function getAdminState_(token) {
    if (!requireAdmin_(token)) { return _apiR(false, null, "Unauthorized"); }


    try {
        const configMap = _readC();

        let stats = null;
        if (configMap['Current Phase'] === 'VACATION_SENIORITY' || configMap['Current Phase'] === 'VACATION_RANDOM') {
            const _vAvail = typeof getWeekAvailability_ === 'function' ? getWeekAvailability_ : (typeof global !== 'undefined' && global.getWeekAvailability_ ? global.getWeekAvailability_ : (require('./Vacation.gs').getWeekAvailability));
            const _vOpts = typeof getAdminOptions_ === 'function' ? getAdminOptions_ : (typeof global !== 'undefined' && global.getAdminOptions_ ? global.getAdminOptions_ : (require('./Vacation.gs').getAdminOptions));
            const _vRoster = typeof getRosterForVacation_ === 'function' ? getRosterForVacation_ : (typeof global !== 'undefined' && global.getRosterForVacation_ ? global.getRosterForVacation_ : (require('./Vacation.gs').getRosterForVacation));

            const adminOpts = _vOpts();
            const roster = _vRoster(configMap['Active Year'], adminOpts.target);
            const { data, map } = _vAvail();

            const weeks = [];
            for (let r = 1; r < data.length; r++) {
                 const row = data[r];
                 const weekId = String(row[map['Vacation Week']] || '').trim();
                 if (!weekId) continue;

                 const capacity = row[map['Capacity Override']] !== '' ? Number(row[map['Capacity Override']]) : adminOpts.capacity;
                 let assignedCount = 0;

                 for (const key of Object.keys(map)) {
                     if (key.startsWith('Person')) {
                         const val = String(row[map[key]] || '').trim();
                         if (val !== '') {
                             assignedCount++;
                         }
                     }
                 }

                 weeks.push({
                     weekId: weekId,
                     type: row[map['Prime Classification']] || 'Non-Prime',
                     capacity: capacity,
                     remaining: capacity - assignedCount
                 });
            }

            // Calculate assigned weeks per roster member
            const rosterStats = roster.map(p => {
                let count = 0;
                for (let r = 1; r < data.length; r++) {
                    for (const key of Object.keys(map)) {
                        if (key.startsWith('Person')) {
                            if (String(data[r][map[key]] || '').trim() === p.name) {
                                count++;
                            }
                        }
                    }
                }
                return {
                    name: p.name,
                    target: p.target,
                    count: count,
                    status: count >= p.target ? 'Complete' : 'Pending'
                };
            });

            const activeTurns = JSON.parse(configMap['Current Active Window'] || '[]');
            const activeNames = activeTurns.map(t => {
                const found = roster.find(r => r.participantId === t.participantId);
                return found ? found.name : t.participantId;
            });

            stats = {
                activeWindow: activeNames,
                roster: rosterStats,
                weeks: weeks
            };
        }

        return _apiR(true, { config: configMap, stats: stats }, 'State retrieved');
    } catch(e) {
        return _apiR(false, null, `State error: ${e.message}`);
    }
}

/**
 * Runs the schema initializer/updater
 * @param {string} token
 * @returns {object} API response containing the report
 */
function runAdminInit_(token) {
    if (!requireAdmin_(token)) { return _apiR(false, null, "Unauthorized"); }


    const result = _initW();
    return result;
}

function beginVacationRound1_(token) {
    if (!requireAdmin_(token)) return _apiR(false, null, "Unauthorized");

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return _apiR(false, null, 'System busy.');

    try {
        const configMap = _readC();
        if (configMap['Phase Ready State'] !== 'READY_VACATION_SENIORITY') {
            return _apiR(false, null, 'Setup is not confirmed. Phase is not ready for vacation.');
        }

        if (configMap['Current Phase'] === 'VACATION_SENIORITY' || configMap['Current Phase'] === 'VACATION_RANDOM') {
            return _apiR(false, null, 'Vacation phase is already active.');
        }

        const _vWrite = typeof writeConfigState_ === 'function' ? writeConfigState_ : (typeof global !== 'undefined' && global.writeConfigState_ ? global.writeConfigState_ : (require('./State.gs').writeConfigState));
        const _vRoster = typeof getRosterForVacation_ === 'function' ? getRosterForVacation_ : (typeof global !== 'undefined' && global.getRosterForVacation_ ? global.getRosterForVacation_ : (require('./Vacation.gs').getRosterForVacation));
        const _vCalcNext = typeof calculateNextQueueState_ === 'function' ? calculateNextQueueState_ : (typeof global !== 'undefined' && global.calculateNextQueueState_ ? global.calculateNextQueueState_ : (require('./QueueEngine.gs').calculateNextQueueState));
        const _vOpts = typeof getAdminOptions_ === 'function' ? getAdminOptions_ : (typeof global !== 'undefined' && global.getAdminOptions_ ? global.getAdminOptions_ : (require('./Vacation.gs').getAdminOptions));

        const opts = _vOpts();
        const roster = _vRoster(configMap['Active Year'], opts.target);

        let newState = { ...configMap, 'Current Phase': 'VACATION_SENIORITY', 'Current Vacation Round': '1' };

        const queueConfig = {
            movementMode: 'FORWARD_ONLY',
            orderSource: 'seniority',
            windowSize: opts.windowSize,
            action: 'INIT',
            phase: 'VACATION_SENIORITY'
        };

        const result = _vCalcNext(newState, roster, queueConfig);

        const updates = {
            "Current Phase": "VACATION_SENIORITY",
            "Current Vacation Round": "1",
            "Current Active Window": JSON.stringify(result.nextState["Current Active Window"]),
            "Current Directional Window": JSON.stringify(result.nextState["Current Directional Window"]),
            "Current Directional Window Completed": JSON.stringify(result.nextState["Current Directional Window Completed"]),
            "Current Queue Skip State": JSON.stringify(result.nextState["Current Queue Skip State"] || {}),
            "Current Queue Cursor": result.nextState["Current Queue Cursor"].toString(),
            "Current Queue Cycle": result.nextState["Current Queue Cycle"].toString(),
            "Current Serpentine Direction": result.nextState["Current Serpentine Direction"],
            "Active Window Generation": result.nextState["Active Window Generation"].toString()
        };

        _vWrite(updates);
        return _apiR(true, null, 'Vacation Round 1 has been started.');
    } catch(e) {
        return _apiR(false, null, 'Error starting vacation round 1: ' + e.message);
    } finally {
        lock.releaseLock();
    }
}

function endVacationEarly_(token) {
    if (!requireAdmin_(token)) return _apiR(false, null, "Unauthorized");

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) return _apiR(false, null, 'System busy.');

    try {
        const configMap = _readC();
        if (configMap['Current Phase'] !== 'VACATION_SENIORITY' && configMap['Current Phase'] !== 'VACATION_RANDOM') {
            return _apiR(false, null, 'Vacation phase is not active.');
        }

        const _vWrite = typeof writeConfigState_ === 'function' ? writeConfigState_ : (typeof global !== 'undefined' && global.writeConfigState_ ? global.writeConfigState_ : (require('./State.gs').writeConfigState));

        const updates = {
            "Phase Ready State": "READY_WEEKEND",
            "Current Active Window": "[]",
            "Current Directional Window": "[]",
            "Current Directional Window Completed": "[]",
            "Current Queue Cursor": "0",
            "Current Queue Cycle": "0",
            "Current Queue Skip State": "{}"
        };

        _vWrite(updates);
        return _apiR(true, null, 'Vacation phase has been ended early. System is ready for Weekend phase.');
    } catch(e) {
        return _apiR(false, null, 'Error ending vacation early: ' + e.message);
    } finally {
        lock.releaseLock();
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        getAdminState: getAdminState_,
        runAdminInit: runAdminInit_,
        requireAdmin_,
        beginVacationRound1: beginVacationRound1_,
        endVacationEarly: endVacationEarly_
    };
}
