/* ───────── 🔧 Helpers & Utilities ───────── */

/**
 * Escapes special regex characters in a string.
 * @param {string} text The text to escape.
 * @returns {string} The escaped text.
 */
function escapeRegex(text) {
  if (!text) return '';
  return text.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

/**
 * Checks if the OpenAI API key is set in Document Properties.
 * @returns {boolean} True if the key exists.
 */
function checkOpenAIKey() {
  try {
    // Сначала попробуем мигрировать ключ, если необходимо
    migrateApiKeyIfNeeded();
    
    // Сначала проверяем Document Properties (новый способ)
    let key = PropertiesService.getDocumentProperties().getProperty('OPENAI_API_KEY');
    let source = 'Document Properties';
    
    // Если не найден, проверяем Script Properties (старый способ)
    if (!key) {
      key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
      source = 'Script Properties';
    }
    
    const hasKey = !!key;
    logOperation('INFO', `🔑 API key check: ${hasKey ? 'Available' : 'Not available'} (from ${source})`, {
      hasKey: hasKey,
      keyLength: key ? key.length : 0,
      keyPreview: key ? key.substring(0, 8) + '***' : 'none',
      source: source
    });
    return hasKey;
  } catch (e) {
    logOperation('ERROR', `🔑 Error checking API key: ${e.message}`, {
      error: e.message,
      stack: e.stack
    });
    return false;
  }
}

/**
 * Устанавливает OpenAI API ключ в Document Properties.
 * @param {string} apiKey API ключ OpenAI.
 * @returns {boolean} True если ключ успешно установлен.
 */
function setOpenAIKey(apiKey) {
  try {
    if (!apiKey || typeof apiKey !== 'string') {
      logOperation('ERROR', '🔑 Invalid API key provided');
      return false;
    }
    
    if (!apiKey.startsWith('sk-')) {
      logOperation('ERROR', '🔑 Invalid API key format (should start with sk-)');
      return false;
    }
    
    // Сохраняем в оба места для совместимости
    PropertiesService.getDocumentProperties().setProperty('OPENAI_API_KEY', apiKey);
    PropertiesService.getScriptProperties().setProperty('OPENAI_API_KEY', apiKey);
    
    logOperation('INFO', `🔑 OpenAI API key успешно установлен (Document + Script Properties)`, {
      keyLength: apiKey.length,
      keyPreview: apiKey.substring(0, 8) + '***'
    });
    return true;
  } catch (e) {
    logOperation('ERROR', `🔑 Error setting API key: ${e.message}`, {
      error: e.message,
      stack: e.stack
    });
    return false;
  }
}

/**
 * Удаляет OpenAI API ключ из Document Properties.
 * @returns {boolean} True если ключ успешно удален.
 */
function clearOpenAIKey() {
  try {
    // Удаляем из обоих мест
    PropertiesService.getDocumentProperties().deleteProperty('OPENAI_API_KEY');
    PropertiesService.getScriptProperties().deleteProperty('OPENAI_API_KEY');
    
    logOperation('INFO', '🔑 OpenAI API key успешно удален (Document + Script Properties)');
    return true;
  } catch (e) {
    logOperation('ERROR', `🔑 Error clearing API key: ${e.message}`, {
      error: e.message,
      stack: e.stack
    });
    return false;
  }
}

/**
 * Мигрирует API ключ из Script Properties в Document Properties для совместимости.
 * Автоматически вызывается при проверке ключа.
 */
function migrateApiKeyIfNeeded() {
  try {
    const documentKey = PropertiesService.getDocumentProperties().getProperty('OPENAI_API_KEY');
    const scriptKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    // Если ключ есть только в Script Properties, копируем его в Document Properties
    if (!documentKey && scriptKey) {
      PropertiesService.getDocumentProperties().setProperty('OPENAI_API_KEY', scriptKey);
      logOperation('INFO', '🔄 API key найден в Script Properties и автоматически мигрирован в Document Properties для совместимости');
      return true;
    }
    
    return false;
  } catch (e) {
    logOperation('ERROR', `🔑 Error migrating API key: ${e.message}`);
    return false;
  }
}

/**
 * Тестирует OpenAI API ключ, делая простой запрос.
 * @returns {boolean} True если ключ работает.
 */
function testOpenAIKey() {
  try {
    // Сначала проверяем Document Properties (новый способ)
    let key = PropertiesService.getDocumentProperties().getProperty('OPENAI_API_KEY');
    
    // Если не найден, проверяем Script Properties (старый способ)
    if (!key) {
      key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    }
    
    if (!key) {
      logOperation('ERROR', '🔑 No API key found for testing');
      return false;
    }
    
    const payload = {
      model: 'gpt-4o-mini',
      messages: [{ role: 'user', content: 'Тест' }],
      max_tokens: 5
    };
    
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    
    const responseCode = response.getResponseCode();
    const isValid = responseCode === 200;
    
    logOperation('INFO', `🔑 API key test: ${isValid ? 'SUCCESS' : 'FAILED'}`, {
      responseCode: responseCode,
      isValid: isValid
    });
    
    return isValid;
  } catch (e) {
    logOperation('ERROR', `🔑 Error testing API key: ${e.message}`, {
      error: e.message,
      stack: e.stack
    });
    return false;
  }
}

/**
 * Generates a unique ID for an element based on its content and position.
 * @param {Object} elementWrapper The element wrapper with text and index.
 * @returns {string} A unique element ID.
 */
function generateElementId(elementWrapper) {
  try {
    const text = elementWrapper.text || '';
    const index = elementWrapper.originalIndex || 0;
    const type = elementWrapper.typeName || 'Unknown';
    
    // Use a simple but effective hashing approach
    const textSample = text.substring(0, 50).replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
    let hash = 0;
    for (let i = 0; i < text.length; i++) {
      const char = text.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32bit integer
    }
    
    return `${type}_${index}_${Math.abs(hash).toString(36)}_${textSample}`;
  } catch (e) {
    logOperation('ERROR', `💥 Error generating element ID: ${e.message}`, {
      error: e.message,
      elementWrapper: elementWrapper
    });
    // Fallback ID
    return `unknown_${elementWrapper.originalIndex || 0}_${Date.now()}`;
  }
}

/**
 * Advanced text normalization for better text matching.
 * @param {string} text The input text.
 * @returns {string} The normalized text.
 */
function normalizeText(text) {
  if (!text) return '';
  return text
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ') // Collapse whitespace
    // Normalize ALL apostrophe variants to a single standard (regular apostrophe)
    .replace(/[\u2019\u2018\u0027\u0060\u02BC\u02CA\u02CB\u0301\u0300\u1FBD\u1FBE\u1FBF\u1FC0\u1FC1\u1FCD\u1FCE\u1FCF\u1FDD\u1FDE\u1FDF\u1FED\u1FEE\u1FEF\u1FFD\u1FFE\u02B9\u02BB\u02BD\u02BE\u02BF\u02C8\u02CC\u02D0\u02D1\u02D2\u02D3\u02D4\u02D5\u02D6\u02D7\u02DE\u02DF\u02E0\u02E1\u02E2\u02E3\u02E4\u0374\u0375\u037A\u0384\u0385]/g, "'") 
    // Normalize dashes to regular hyphen
    .replace(/[–—−]/g, '-') 
    .trim();
}

/**
 * Gets a human-readable name for an element type.
 * @param {GoogleAppsScript.Document.ElementType} elementType The element type enum.
 * @returns {string} The readable name.
 */
function getElementTypeName(elementType) {
  // Using a simple map for clarity and performance
  const typeMap = {
    [DocumentApp.ElementType.PARAGRAPH]: 'Paragraph',
    [DocumentApp.ElementType.HEADING]: 'Heading', 
    [DocumentApp.ElementType.LIST_ITEM]: 'List Item',
    [DocumentApp.ElementType.TABLE]: 'Table',
    [DocumentApp.ElementType.TABLE_CELL]: 'Table Cell',
    [DocumentApp.ElementType.TEXT]: 'Text',
    [DocumentApp.ElementType.BODY_SECTION]: 'Body'
  };
  return typeMap[elementType] || 'Unsupported';
}
