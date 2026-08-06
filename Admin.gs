// Injection wrappers
const _apiR = typeof apiResponse === 'function' ? apiResponse : (typeof global !== 'undefined' && global.apiResponse ? global.apiResponse : (require('./Utils.gs').apiResponse));
const _readC = typeof readConfigState === 'function' ? readConfigState : (typeof global !== 'undefined' && global.readConfigState ? global.readConfigState : (require('./State.gs').readConfigState));
const _initW = typeof initializeOrUpdateWorkbook === 'function' ? initializeOrUpdateWorkbook : (typeof global !== 'undefined' && global.initializeOrUpdateWorkbook ? global.initializeOrUpdateWorkbook : (require('./Schema.gs').initializeOrUpdateWorkbook));

/**
 * Validates the admin token before proceeding with any action
 * @param {string} token
 * @returns {boolean}
 */
function requireAdmin_(token) {
    const _resAdmin = typeof resolveAdminSession === 'function' ? resolveAdminSession : (typeof global !== 'undefined' && global.resolveAdminSession ? global.resolveAdminSession : (require('./Auth.gs').resolveAdminSession));
    return _resAdmin(token);
}

/**
 * Returns safe durable-state values for the dashboard
 * @param {string} token
 * @returns {object} API response
 */
function getAdminState(token) {
    if (!requireAdmin_(token)) {
        return _apiR(false, null, 'Unauthorized');
    }

    try {
        const configMap = _readC();
        return _apiR(true, configMap, 'State retrieved');
    } catch(e) {
        return _apiR(false, null, `State error: ${e.message}`);
    }
}

/**
 * Runs the schema initializer/updater
 * @param {string} token
 * @returns {object} API response containing the report
 */
function runAdminInit(token) {
    if (!requireAdmin_(token)) {
        return _apiR(false, null, 'Unauthorized');
    }

    const result = _initW();
    return result;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        getAdminState,
        runAdminInit,
        requireAdmin_
    };
}
