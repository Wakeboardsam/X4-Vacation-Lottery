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
