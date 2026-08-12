// Dependencies
const _rState = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));
const _rApi = typeof apiResponse_ === 'function' ? apiResponse_ : (typeof global !== 'undefined' && global.apiResponse_ ? global.apiResponse_ : (require('./Utils.gs').apiResponse));
const _rHMap = typeof getHeaderMap_ === 'function' ? getHeaderMap_ : (typeof global !== 'undefined' && global.getHeaderMap_ ? global.getHeaderMap_ : (require('./Utils.gs').getHeaderMap));
const _rReqPart = () => { return typeof resolveParticipantSession_ === 'function' ? resolveParticipantSession_ : (typeof global !== 'undefined' && global.resolveParticipantSession_ ? global.resolveParticipantSession_ : (require('./Auth.gs').resolveParticipantSession)); };

function getRulesContent_() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Rules & Tips');
    if (!sheet) {
        return {
            rules: [],
            remindersByContext: { VACATION: [], WEEKEND: [], HOLIDAY: [], TRANSFER: [], SPOUSE: [] },
            spouseReminder: ""
        };
    }

    const data = sheet.getDataRange().getValues();
    const map = _rHMap(data);

    let rules = [];
    let remindersByContext = { VACATION: [], WEEKEND: [], HOLIDAY: [], TRANSFER: [], SPOUSE: [] };
    let spouseReminder = "";

    // Backward compatibility headers
    const colOrder = map['Display Order'];
    const colText = map['Rule Text'];
    const colEnabled = map['Enabled'];
    const colKey = map['Content Key'];
    const colType = map['Content Type'];
    const colContext = map['Context'];

    if (colOrder === undefined || colText === undefined || colEnabled === undefined) {
         return {
            rules: [],
            remindersByContext: { VACATION: [], WEEKEND: [], HOLIDAY: [], TRANSFER: [], SPOUSE: [] },
            spouseReminder: ""
         };
    }

    for (let r = 1; r < data.length; r++) {
        const row = data[r];
        const enabled = String(row[colEnabled] || '').toLowerCase() === 'true';
        if (!enabled) continue;

        const text = String(row[colText] || '').trim();
        if (!text) continue;

        const order = Number(row[colOrder]) || 999;
        const type = (colType !== undefined && String(row[colType] || '').trim()) ? String(row[colType] || '').trim() : 'RULE';
        const context = (colContext !== undefined && String(row[colContext] || '').trim()) ? String(row[colContext] || '').trim() : 'GENERAL';

        if (type === 'RULE') {
            rules.push({ text: text, context: context, displayOrder: order, originalRowIndex: r });
        } else if (type === 'REMINDER') {
            if (remindersByContext[context]) {
                remindersByContext[context].push(text);
            }
        } else if (type === 'SPOUSE_REMINDER') {
             spouseReminder = text; // Just take the first or last enabled one?
             if (remindersByContext['SPOUSE']) {
                  remindersByContext['SPOUSE'].push(text);
             }
        }
    }

    // Sort rules by display order, then by row index (stable sort for ties)
    rules.sort((a, b) => {
        if (a.displayOrder !== b.displayOrder) {
            return a.displayOrder - b.displayOrder;
        }
        return a.originalRowIndex - b.originalRowIndex;
    });

    // Remove internal fields for client
    const clientRules = rules.map(r => ({ text: r.text, context: r.context, displayOrder: r.displayOrder }));

    return {
        rules: clientRules,
        remindersByContext: remindersByContext,
        spouseReminder: spouseReminder || (remindersByContext['SPOUSE'] && remindersByContext['SPOUSE'].length > 0 ? remindersByContext['SPOUSE'][0] : "")
    };
}

function mapTransferPreferences_(transferChoice) {
    let giver = false;
    let receiver = false;

    if (transferChoice === 'OFFER') {
        giver = true;
    } else if (transferChoice === 'RECEIVE') {
        receiver = true;
    } else if (transferChoice === 'BOTH') {
        giver = true;
        receiver = true;
    } else if (transferChoice === 'NONE') {
        // Both false
    } else {
        return null;
    }

    return { giver, receiver };
}

function getParticipantAcknowledgmentState_(participantId, activeYear) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Turn Management');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const map = _rHMap(data);

    if (map['Participant ID'] === undefined || map['Rules Acknowledged Year'] === undefined) return null;

    for (let r = 1; r < data.length; r++) {
         if (String(data[r][map['Participant ID']] || '').trim() === participantId) {
             const ackYear = String(data[r][map['Rules Acknowledged Year']] || '').trim();
             return {
                 acknowledgedYear: ackYear,
                 acknowledged: ackYear === String(activeYear).trim() && ackYear !== '',
                 holidayVolunteer: Boolean(data[r][map['Holiday Volunteer']]),
                 transferGiver: Boolean(data[r][map['Transfer Giver']]),
                 transferReceiver: Boolean(data[r][map['Transfer Receiver']])
             };
         }
    }
    return null;
}

function isParticipantAcknowledgedForYear_(participantId, activeYear) {
    const state = getParticipantAcknowledgmentState_(participantId, activeYear);
    return state ? state.acknowledged : false;
}

function saveRulesAcknowledgment_(token, holidayVolunteerAnswer, transferPreference) {
    const proj = _rReqPart()(token);
    if (!proj) return _rApi(false, null, 'Session expired or invalid');

    if (holidayVolunteerAnswer !== 'YES' && holidayVolunteerAnswer !== 'NO') {
        return _rApi(false, null, 'Choose Yes or No for Holiday Volunteer and select a transfer preference before continuing.');
    }

    const transferMapped = mapTransferPreferences_(transferPreference);
    if (!transferMapped) {
        return _rApi(false, null, 'Choose Yes or No for Holiday Volunteer and select a transfer preference before continuing.');
    }

    const holidayVolunteer = holidayVolunteerAnswer === 'YES';

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(10000)) {
        return _rApi(false, null, 'System is busy processing another request.');
    }

    try {
        const configMap = _rState();
        const activeYear = String(configMap['Active Year'] || '').trim();
        if (!activeYear || activeYear.length !== 4) return _rApi(false, null, 'Active Year is not configured.');

        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('Turn Management');
        if (!sheet) return _rApi(false, null, 'Turn Management not found.');

        const data = sheet.getDataRange().getValues();
        const map = _rHMap(data);

        let matchRowIdx = -1;
        let matchCount = 0;
        for (let r = 1; r < data.length; r++) {
             if (String(data[r][map['Participant ID']] || '').trim() === proj.participantId) {
                 matchRowIdx = r;
                 matchCount++;
             }
        }

        if (matchCount !== 1) {
             return _rApi(false, null, 'Participant record not found or duplicate.');
        }

        const currentAckYear = String(data[matchRowIdx][map['Rules Acknowledged Year']] || '').trim();
        if (currentAckYear === activeYear && currentAckYear !== '') {
            // Already acknowledged, return success without rewriting
            return _rApi(true, null, 'Rules already acknowledged for the active year.');
        }

        // Snapshot values for rollback
        const origHoliday = data[matchRowIdx][map['Holiday Volunteer']];
        const origGiver = data[matchRowIdx][map['Transfer Giver']];
        const origReceiver = data[matchRowIdx][map['Transfer Receiver']];
        const origAckYear = data[matchRowIdx][map['Rules Acknowledged Year']];

        try {
            // Write atomically as batch
            sheet.getRange(matchRowIdx + 1, map['Holiday Volunteer'] + 1).setValue(holidayVolunteer);
            sheet.getRange(matchRowIdx + 1, map['Transfer Giver'] + 1).setValue(transferMapped.giver);
            sheet.getRange(matchRowIdx + 1, map['Transfer Receiver'] + 1).setValue(transferMapped.receiver);
            sheet.getRange(matchRowIdx + 1, map['Rules Acknowledged Year'] + 1).setValue(activeYear);
        } catch (writeErr) {
            // Rollback
            try {
                sheet.getRange(matchRowIdx + 1, map['Holiday Volunteer'] + 1).setValue(origHoliday);
                sheet.getRange(matchRowIdx + 1, map['Transfer Giver'] + 1).setValue(origGiver);
                sheet.getRange(matchRowIdx + 1, map['Transfer Receiver'] + 1).setValue(origReceiver);
                sheet.getRange(matchRowIdx + 1, map['Rules Acknowledged Year'] + 1).setValue(origAckYear);
            } catch (rollbackErr) {
                throw new Error(`Fatal Consistency Error: Acknowledgment save failed (${writeErr.message}), AND compensating rollback failed (${rollbackErr.message}).`);
            }
            return _rApi(false, null, 'Failed to save rules acknowledgment.');
        }

        return _rApi(true, null, 'Rules acknowledged successfully.');

    } catch (e) {
        if (e.message.includes('Fatal Consistency Error')) {
            throw e;
        }
        return _rApi(false, null, 'An internal error occurred: ' + e.message);
    } finally {
        lock.releaseLock();
    }
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        getRulesContent: getRulesContent_,
        mapTransferPreferences: mapTransferPreferences_,
        getParticipantAcknowledgmentState: getParticipantAcknowledgmentState_,
        isParticipantAcknowledgedForYear: isParticipantAcknowledgedForYear_,
        saveRulesAcknowledgment: saveRulesAcknowledgment_
    };
}