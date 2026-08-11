const ADMIN_SALT_KEY = 'ADMIN_SALT';
const ADMIN_HASH_KEY = 'ADMIN_HASH';
const SESSION_PREFIX_USER = 'SESS_U_';
const SESSION_PREFIX_ADMIN = 'SESS_A_';
const SESSION_EXPIRY_MS = 6 * 60 * 60 * 1000; // 6 hours

// Dependency injection workaround for Apps Script sharing environment vs Node tests
const _apiResponse = typeof apiResponse_ === 'function' ? apiResponse_ : (typeof global !== 'undefined' && global.apiResponse_ ? global.apiResponse_ : (require('./Utils.gs').apiResponse));
const _getHMap2 = typeof getHeaderMap_ === 'function' ? getHeaderMap_ : (typeof global !== 'undefined' && global.getHeaderMap_ ? global.getHeaderMap_ : (require('./Utils.gs').getHeaderMap));
const _validateParticipantIds2 = typeof validateParticipantIds_ === 'function' ? validateParticipantIds_ : (typeof global !== 'undefined' && global.validateParticipantIds_ ? global.validateParticipantIds_ : (require('./Utils.gs').validateParticipantIds));

/**
 * Editor-only function to set the administrator access code.
 * @param {string} rawCode
 */
function setAdminAccessCode_(rawCode) {
    if (!rawCode || typeof rawCode !== 'string' || rawCode.trim() === '') {
        throw new Error('Access code cannot be blank.');
    }
    const salt = Utilities.getUuid();
    const salted = salt + rawCode;
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salted, Utilities.Charset.UTF_8);
    const hashHex = digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

    PropertiesService.getScriptProperties().setProperties({
        [ADMIN_SALT_KEY]: salt,
        [ADMIN_HASH_KEY]: hashHex
    });
}

function cleanExpiredSessions_(props, prefix) {
    const all = props.getProperties();
    const now = Date.now();
    for (const k in all) {
        if (k.startsWith(prefix)) {
            try {
                const parsed = JSON.parse(all[k]);
                if (now > parsed.expires) {
                    props.deleteProperty(k);
                }
            } catch (e) {
                props.deleteProperty(k);
            }
        }
    }
}

/**
 * Authenticates participant using PIN.
 * @param {string} pin
 * @returns {object} API response with session token and projection
 */
function loginParticipant_(pin) {
    if (!pin || String(pin).trim() === '') {
         return _apiResponse(false, null, 'Invalid PIN'); // generic error
    }

    const pinStr = String(pin).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Turn Management');
    if (!sheet) {
        return _apiResponse(false, null, 'System not initialized');
    }

    const data = sheet.getDataRange().getValues();
    const map = _getHMap2(data);

    try {
        _validateParticipantIds2(data, map);
    } catch (identityErr) {
        return _apiResponse(false, null, identityErr.participantMessage || 'Vacation selection is temporarily unavailable because the participant roster requires administrator correction. No changes were made.');
    }

    if (map['PIN'] === undefined) {
         return _apiResponse(false, null, 'System missing PIN configuration');
    }

    let matchIdx = -1;
    let duplicateFound = false;

    for (let r = 1; r < data.length; r++) {
         const rowPin = String(data[r][map['PIN']] || '').trim();
         if (rowPin === pinStr) {
             if (matchIdx !== -1) {
                 duplicateFound = true;
             } else {
                 matchIdx = r;
             }
         }
    }

    if (duplicateFound) {
        // Log to admin if we had logging, but for participant, generic fail
        return _apiResponse(false, null, 'Invalid PIN');
    }
    if (matchIdx === -1) {
        return _apiResponse(false, null, 'Invalid PIN');
    }

    const participantName = data[matchIdx][map['Name']];
    const participantId = String(data[matchIdx][map['Participant ID']] || '').trim();
    const activeForYear = Boolean(data[matchIdx][map['Active for Year']]);

    // Read high level state
    let configMap;
    try {
        const _readConfig = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));
        configMap = _readConfig();
    } catch(e) {
        return _apiResponse(false, null, 'System error reading state');
    }

    // Create session
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) {
        return _apiResponse(false, null, 'System busy, try again later');
    }

    let rawToken;
    try {
        const props = PropertiesService.getScriptProperties();
        cleanExpiredSessions_(props, SESSION_PREFIX_USER);

        rawToken = Utilities.getUuid();
        const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawToken, Utilities.Charset.UTF_8);
        const tokenHashHex = digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

        const sessionData = {
            expires: Date.now() + SESSION_EXPIRY_MS,
            participantId: participantId
        };

        props.setProperty(SESSION_PREFIX_USER + tokenHashHex, JSON.stringify(sessionData));
    } finally {
        lock.releaseLock();
    }

    return _apiResponse(true, {
        token: rawToken,
        participant: {
            participantId: participantId,
            name: participantName,
            activeYear: configMap['Active Year'],
            setupState: configMap['Setup State'],
            currentPhase: configMap['Current Phase'],
            phaseReadyState: configMap['Phase Ready State'],
            isActive: activeForYear
        }
    }, 'Login successful');
}

/**
 * Resolves a participant session token to their projection.
 * @param {string} token
 * @returns {object} projection or null if invalid
 */
function resolveParticipantSession_(token) {
    if (!token) return null;

    const digestForResolve = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8);
    const hexForResolve = digestForResolve.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

    const props = PropertiesService.getScriptProperties();
    const sessStr = props.getProperty(SESSION_PREFIX_USER + hexForResolve);

    if (!sessStr) return null;

    let sess;
    try {
        sess = JSON.parse(sessStr);
    } catch(e) {
        return null;
    }

    if (Date.now() > sess.expires) {
         props.deleteProperty(SESSION_PREFIX_USER + hexForResolve);
         return null;
    }

    // Reread participant record to ensure fresh state
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Turn Management');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    const map = _getHMap2(data);

    try {
        _validateParticipantIds2(data, map);
    } catch (identityErr) {
        props.deleteProperty(SESSION_PREFIX_USER + hexForResolve);
        throw identityErr;
    }

    let matchRow = null;
    let duplicateCount = 0;
    for (let r = 1; r < data.length; r++) {
         if (String(data[r][map['Participant ID']] || '').trim() === String(sess.participantId || '').trim()) {
              matchRow = data[r];
              duplicateCount++;
         }
    }

    if (duplicateCount > 1) {
        props.deleteProperty(SESSION_PREFIX_USER + hexForResolve);
        return null;
    }

    if (!matchRow) return null; // Participant deleted or PIN changed

    let configMap;
    try {
        const _readConfig = typeof readConfigState_ === 'function' ? readConfigState_ : (typeof global !== 'undefined' && global.readConfigState_ ? global.readConfigState_ : (require('./State.gs').readConfigState));
        configMap = _readConfig();
    } catch(e) {
        return null;
    }

    return {
        participantId: String(matchRow[map['Participant ID']] || '').trim(),
        name: matchRow[map['Name']],
        activeYear: configMap['Active Year'],
        setupState: configMap['Setup State'],
        currentPhase: configMap['Current Phase'],
        phaseReadyState: configMap['Phase Ready State'],
        isActive: Boolean(matchRow[map['Active for Year']])
    };
}


/**
 * Authenticates Administrator.
 * @param {string} rawCode
 * @returns {object} API response with admin session token
 */
function loginAdmin_(rawCode) {
    if (!rawCode) return _apiResponse(false, null, 'Invalid code');

    const props = PropertiesService.getScriptProperties();
    const salt = props.getProperty(ADMIN_SALT_KEY);
    const expectedHash = props.getProperty(ADMIN_HASH_KEY);

    if (!salt || !expectedHash) {
         return _apiResponse(false, null, 'Admin access not configured');
    }

    const salted = salt + rawCode;
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salted, Utilities.Charset.UTF_8);
    const hashHex = digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

    // Timing safe compare could be implemented here, string equality is standard for basic GAS if lengths match
    if (hashHex !== expectedHash) {
        return _apiResponse(false, null, 'Invalid code');
    }

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return _apiResponse(false, null, 'System busy');

    let rawToken;
    try {
        cleanExpiredSessions_(props, SESSION_PREFIX_ADMIN);
        rawToken = Utilities.getUuid();
        const tDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, rawToken, Utilities.Charset.UTF_8);
        const tHashHex = tDigest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

        props.setProperty(SESSION_PREFIX_ADMIN + tHashHex, JSON.stringify({ expires: Date.now() + SESSION_EXPIRY_MS }));
    } finally {
        lock.releaseLock();
    }

    return _apiResponse(true, { token: rawToken }, 'Admin login successful');
}


function resolveAdminSession_(token) {
    if (!token) return false;

    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8);
    const tokenHashHex = digest.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

    const props = PropertiesService.getScriptProperties();
    const digest3 = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8);
    const hex3 = digest3.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');
    const sessStr = props.getProperty(SESSION_PREFIX_ADMIN + hex3);

    if (!sessStr) return false;

    let sess;
    try {
        sess = JSON.parse(sessStr);
    } catch(e) {
        return false;
    }

    if (Date.now() > sess.expires) {
         props.deleteProperty(SESSION_PREFIX_ADMIN + hex3);
         return false;
    }
    return true;
}



function logout_(token) {
    if (!token) return _apiResponse(true, null, 'Logged out');

    // Scoped block to avoid variable shadowing
    var _dgst = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token, Utilities.Charset.UTF_8);
    var _h = _dgst.map(b => (b < 0 ? b + 256 : b).toString(16).padStart(2, '0')).join('');

    const lock = LockService.getScriptLock();
    if (lock.tryLock(3000)) {
        try {
            const props = PropertiesService.getScriptProperties();
            let success = false;

            if (props.getProperty(SESSION_PREFIX_USER + _h)) {
                props.deleteProperty(SESSION_PREFIX_USER + _h);
                success = true;
            } else if (props.getProperty(SESSION_PREFIX_ADMIN + _h)) {
                props.deleteProperty(SESSION_PREFIX_ADMIN + _h);
                success = true;
            }
            if (!success) {
                 return _apiResponse(false, null, 'Session not found');
            }
        } catch(e) {
             return _apiResponse(false, null, 'Error during logout');
        } finally {
            lock.releaseLock();
        }
    } else {
         return _apiResponse(false, null, 'System busy');
    }
    return _apiResponse(true, null, 'Logged out');
}



function requireAdmin_(token) {
    return resolveAdminSession_(token);
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        setAdminAccessCode: setAdminAccessCode_,
        loginParticipant: loginParticipant_,
        resolveParticipantSession: resolveParticipantSession_,
        loginAdmin: loginAdmin_,
        resolveAdminSession: resolveAdminSession_,
        logout: logout_,
        requireAdmin: requireAdmin_,
        SESSION_PREFIX_USER,
        SESSION_PREFIX_ADMIN
    };
}
