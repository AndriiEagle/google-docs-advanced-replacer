/**
 * Global variables for managing addon state.
 * In Google Apps Script, all global variables across all .gs files share the same scope.
 */

// Use conditional initialization to avoid re-declaration errors if legacy files define the same vars
if (typeof OPERATION_LOGS === 'undefined') {
  var OPERATION_LOGS = [];
}

/**
 * Global AI cache to avoid repeated requests for the same data.
 */
if (typeof AI_SEARCH_CACHE === 'undefined') {
  var AI_SEARCH_CACHE = new Map();
}
