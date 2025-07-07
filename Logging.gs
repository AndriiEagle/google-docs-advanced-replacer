/* ───────── 📊 Logging & Error Reporting ───────── */

/**
 * Enhanced logging function that collects logs for later retrieval.
 * Logs are stored in the global OPERATION_LOGS array.
 * @param {string} level Log level (INFO, WARN, ERROR)
 * @param {string} message Log message
 * @param {Object} [data=null] Additional data to log
 */
function logOperation(level, message, data = null) {
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp: timestamp,
    level: level,
    message: message,
    data: data
  };
  
  OPERATION_LOGS.push(logEntry);
  
  try {
    // Persist to Document Cache (up to 100 KB) for cross-execution retrieval
    const cache = CacheService.getDocumentCache();
    cache.put('LAST_OPERATION_LOGS', JSON.stringify(OPERATION_LOGS), 21600); // 6 h
  } catch (e) {
    console.warn('⚠️ Unable to persist logs to CacheService:', e.message);
  }
  try {
    // Small fallback copy (<9 KB) to Document Properties (legacy)
    const json = JSON.stringify(OPERATION_LOGS);
    if (json.length < 9000) {
      PropertiesService.getDocumentProperties().setProperty('LAST_OPERATION_LOGS', json);
    }
  } catch (e) {
    console.warn('⚠️ Unable to persist logs to Document Properties:', e.message);
  }
  
  // Also log to the Apps Script console with appropriate level
  const fullMessage = `[${timestamp}] ${level}: ${message}`;
  switch (level) {
    case 'ERROR':
      console.error(fullMessage, data || '');
      break;
    case 'WARN':
      console.warn(fullMessage, data || '');
      break;
    default:
      console.log(fullMessage, data || '');
  }
  
  // Keep only the last 1000 log entries to prevent memory issues
  if (OPERATION_LOGS.length > 1000) {
    OPERATION_LOGS.splice(0, OPERATION_LOGS.length - 1000);
  }
}

/**
 * Clears the operation logs.
 */
function clearOperationLogs() {
  OPERATION_LOGS.length = 0;
  try {
    PropertiesService.getDocumentProperties().deleteProperty('LAST_OPERATION_LOGS');
  } catch (e) {
    console.warn('⚠️ Unable to clear persisted logs:', e.message);
  }
  try {
    CacheService.getDocumentCache().remove('LAST_OPERATION_LOGS');
  } catch (e) {
    console.warn('⚠️ Unable to clear cached logs:', e.message);
  }
  logOperation('INFO', '🗑️ Operation logs cleared');
}

/**
 * Gets all collected logs for the current operation.
 * @returns {Array<Object>} A copy of the log entries array.
 */
function getOperationLogs() {
  if (OPERATION_LOGS.length > 0) {
    return [...OPERATION_LOGS];
  }
  // If no logs in current execution context, attempt to load from CacheService then Document Properties
  try {
    const cache = CacheService.getDocumentCache();
    const cached = cache.get('LAST_OPERATION_LOGS');
    if (cached) {
      return JSON.parse(cached);
    }
  } catch (e) {
    console.warn('⚠️ Unable to load cached logs:', e.message);
  }
  // Fallback to Document Properties
  try {
    const stored = PropertiesService.getDocumentProperties().getProperty('LAST_OPERATION_LOGS');
    if (stored) {
      return JSON.parse(stored);
    }
  } catch (e) {
    console.warn('⚠️ Unable to load persisted logs:', e.message);
  }
  return [];
}

/**
 * Gets a summary of errors and warnings from the logs.
 * @returns {Object} An object containing counts and messages for logs.
 */
function getLogSummary() {
  const errors = OPERATION_LOGS.filter(log => log.level === 'ERROR');
  const warnings = OPERATION_LOGS.filter(log => log.level === 'WARN');
  
  return {
    totalLogs: OPERATION_LOGS.length,
    errors: errors.length,
    warnings: warnings.length,
    errorMessages: errors.map(e => e.message),
    warningMessages: warnings.map(w => w.message),
    lastLogTime: OPERATION_LOGS.length > 0 ? OPERATION_LOGS[OPERATION_LOGS.length - 1].timestamp : null
  };
}
