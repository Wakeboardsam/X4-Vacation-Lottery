/**
 * Code.gs - Entry point and public wrappers
 */

function doGet(e) {
  // Use template to embed external files later if needed
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('X4 Vacation Lottery')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('X4 Lottery')
      .addItem('Initialize / Update Workbook', 'menuInitWorkbook_')
      .addToUi();
}

/**
 * Menu entry point for manual setup
 */
function menuInitWorkbook_() {
  const result = initializeOrUpdateWorkbook_();
  const ui = SpreadsheetApp.getUi();
  if (result.ok) {
     ui.alert('Success', result.message + '\n\n' + (result.data.changed ? 'Changes were applied.' : 'No changes were needed.'), ui.ButtonSet.OK);
  } else {
     ui.alert('Error', result.message + '\n\nPlease check logs or conflicts.', ui.ButtonSet.OK);
  }
}

// --- Thin API Wrappers for Client ---

function apiLoginParticipant(pin) {
    try {
        return loginParticipant_(pin);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiGetParticipantVacationData(token) {
    try {
        const _reqPart = typeof resolveParticipantSession_ === 'function' ? resolveParticipantSession_ : (typeof global !== 'undefined' && global.resolveParticipantSession_ ? global.resolveParticipantSession_ : (require('./Auth.gs').resolveParticipantSession));
        const proj = _reqPart(token);
        if (!proj) return apiResponse_(false, null, 'Session expired or invalid');

        const _vAvail = typeof getWeekAvailability_ === 'function' ? getWeekAvailability_ : (typeof global !== 'undefined' && global.getWeekAvailability_ ? global.getWeekAvailability_ : (require('./Vacation.gs').getWeekAvailability));
        const _vOpts = typeof getAdminOptions_ === 'function' ? getAdminOptions_ : (typeof global !== 'undefined' && global.getAdminOptions_ ? global.getAdminOptions_ : (require('./Vacation.gs').getAdminOptions));
        const _vRoster = typeof getRosterForVacation_ === 'function' ? getRosterForVacation_ : (typeof global !== 'undefined' && global.getRosterForVacation_ ? global.getRosterForVacation_ : (require('./Vacation.gs').getRosterForVacation));
        const _readC = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));

        const configMap = _readC();
        const activeWindow = JSON.parse(configMap['Current Active Window'] || '[]');
        const turn = activeWindow.find(t => t.participantId === proj.participantId);

        const adminOpts = _vOpts();
        const roster = _vRoster(configMap['Active Year'], adminOpts.target);
        const pRoster = roster.find(r => r.participantId === proj.participantId);

        const { data, map } = _vAvail();
        const weeks = [];
        for (let r = 1; r < data.length; r++) {
             const row = data[r];
             const weekId = String(row[map['Vacation Week']] || '').trim();
             if (!weekId) continue;

             const capacity = row[map['Capacity Override']] !== '' ? Number(row[map['Capacity Override']]) : adminOpts.capacity;
             let assignedCount = 0;
             let assignedToMe = false;

             for (const key of Object.keys(map)) {
                 if (key.startsWith('Person')) {
                     const val = String(row[map[key]] || '').trim();
                     if (val !== '') {
                         assignedCount++;
                         if (val === pRoster.name) assignedToMe = true;
                     }
                 }
             }

             weeks.push({
                 weekId: weekId,
                 primeType: row[map['Prime Classification']] || 'Non-Prime',
                 specialType: row[map['Special Week']] || 'None',
                 capacity: capacity,
                 assignedCount: assignedCount,
                 assignedToMe: assignedToMe
             });
        }

        return apiResponse_(true, {
            turn: turn,
            roster: pRoster,
            weeks: weeks,
            config: {
                 phase: configMap['Current Phase'],
                 round: configMap['Current Vacation Round'],
                 direction: configMap['Current Serpentine Direction']
            }
        }, 'Vacation data retrieved.');
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred: ' + e.message);
    }
}

function apiLoginAdmin(code) {
    try {
        return loginAdmin_(code);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiLogout(token) {
    try {
        return logout_(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiResolveParticipant(token) {
    try {
        const proj = resolveParticipantSession_(token);
        if (proj) return apiResponse_(true, proj, 'Session valid');
        return apiResponse_(false, null, 'Session expired or invalid');
    } catch(e) {
         Logger.log(e);
         return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiGetAdminState(token) {
    try {
        const _reqAdmin = typeof requireAdmin_ === 'function' ? requireAdmin_ : (typeof global !== 'undefined' && global.requireAdmin_ ? global.requireAdmin_ : (require('./Auth.gs').requireAdmin));
        _reqAdmin(token);
        return getAdminState_(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiRunAdminInit(token) {
    try {
        const _reqAdmin = typeof requireAdmin_ === 'function' ? requireAdmin_ : (typeof global !== 'undefined' && global.requireAdmin_ ? global.requireAdmin_ : (require('./Auth.gs').requireAdmin));
        _reqAdmin(token);
        return runAdminInit_(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiBeginVacationRound1(token) {
    try {
        const _reqAdmin = typeof requireAdmin_ === 'function' ? requireAdmin_ : (typeof global !== 'undefined' && global.requireAdmin_ ? global.requireAdmin_ : (require('./Auth.gs').requireAdmin));
        _reqAdmin(token);
        return beginVacationRound1_(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiEndVacationEarly(token) {
    try {
        const _reqAdmin = typeof requireAdmin_ === 'function' ? requireAdmin_ : (typeof global !== 'undefined' && global.requireAdmin_ ? global.requireAdmin_ : (require('./Auth.gs').requireAdmin));
        _reqAdmin(token);
        return endVacationEarly_(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiSubmitVacation(token, turnId, selections) {
    try {
        const _reqPart = typeof resolveParticipantSession_ === 'function' ? resolveParticipantSession_ : (typeof global !== 'undefined' && global.resolveParticipantSession_ ? global.resolveParticipantSession_ : (require('./Auth.gs').resolveParticipantSession));
        const proj = _reqPart(token);
        if (!proj) return apiResponse_(false, null, 'Session expired or invalid');

        return submitVacation_(proj.participantId, turnId, selections);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}

function apiInspectSchema(token) {
    try {
        const _reqAdmin = typeof requireAdmin_ === 'function' ? requireAdmin_ : (typeof global !== 'undefined' && global.requireAdmin_ ? global.requireAdmin_ : (require('./Auth.gs').requireAdmin));
        _reqAdmin(token);
        return inspectSchema_(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse_(false, null, 'An internal error occurred.');
    }
}
