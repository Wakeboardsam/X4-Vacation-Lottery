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
function runAdminInit_(token) {
    if (!requireAdmin_(token)) { return _apiR(false, null, "Unauthorized"); }


    const result = _initW();
    return result;
}

if (typeof module !== 'undefined' && module.exports) {
    module.exports = {
        getAdminState: getAdminState_,
        runAdminInit: runAdminInit_,
        requireAdmin_
    };
}
