/**
 * Code.gs - Entry point and public wrappers
 */

function doGet(e) {
  // Use template to embed external files later if needed
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('X4 Vacation Lottery')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('X4 Lottery')
      .addItem('Initialize / Update Workbook', 'menuInitWorkbook')
      .addToUi();
}

/**
 * Menu entry point for manual setup
 */
function menuInitWorkbook() {
  const result = initializeOrUpdateWorkbook();
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
        return loginParticipant(pin);
    } catch(e) {
        Logger.log(e);
        return apiResponse(false, null, 'An internal error occurred.');
    }
}

function apiLoginAdmin(code) {
    try {
        return loginAdmin(code);
    } catch(e) {
        Logger.log(e);
        return apiResponse(false, null, 'An internal error occurred.');
    }
}

function apiLogout(token, type) {
    try {
        return logout(token, type);
    } catch(e) {
        Logger.log(e);
        return apiResponse(false, null, 'An internal error occurred.');
    }
}

function apiResolveParticipant(token) {
    try {
        const proj = resolveParticipantSession(token);
        if (proj) return apiResponse(true, proj, 'Session valid');
        return apiResponse(false, null, 'Session expired or invalid');
    } catch(e) {
         Logger.log(e);
         return apiResponse(false, null, 'An internal error occurred.');
    }
}

function apiGetAdminState(token) {
    try {
        return getAdminState(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse(false, null, 'An internal error occurred.');
    }
}

function apiRunAdminInit(token) {
    try {
        return runAdminInit(token);
    } catch(e) {
        Logger.log(e);
        return apiResponse(false, null, 'An internal error occurred.');
    }
}
