/*
==========================================================================================
=== Advanced Replacer for Google Docs - Core Logic ===
==========================================================================================
*
* DESCRIPTION: This file contains the main functions that orchestrate the addon's
*              behavior, including UI setup, preview generation, and applying changes.
*
==========================================================================================
*/

/* ───────── UI & Menu ───────── */

/**
 * Creates the Add-on menu in the Google Docs UI when the document is opened.
 * This is a special trigger function recognized by Google Apps Script.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('🚀 Advanced Replacer')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * Displays the HTML sidebar.
 * This function will be updated to use a templated approach for HTML.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('🚀 Advanced Replacer');
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Includes an HTML file content within another. Used for templating.
 * @param {string} filename The name of the HTML file to include.
 * @return {string} The content of the file.
 */
function include(filename) {
  return HtmlService.createTemplateFromFile(filename).evaluate().getContent();
}

/**
 * Displays the UI as a modal dialog for an expanded view.
 */
function showAsModal(directivesText) {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setWidth(1900)
      .setHeight(1500);
 
  htmlOutput.append(`<script>
      window.addEventListener('DOMContentLoaded', function() {
        if (typeof init === 'function') {
            init(${JSON.stringify(directivesText)});
        } else {
            // Fallback if init is not ready
            document.getElementById('directives').value = ${JSON.stringify(directivesText)};
        }
      });
    </script>`);

  DocumentApp.getUi().showModalDialog(htmlOutput, '🚀 Advanced Replacer (Expanded View)');
}

/* ───────── 🚀 Main Processing Engine ───────── */

/**
 * Generates replacement previews based on a single pass over the document.
 * This is the main function called from the client-side JavaScript.
 * @param {Array<Object>} directives An array of {fragment, replaceWith} objects.
 * @param {Object} settings An object with settings like {aiThreshold}.
 * @returns {Array<Object>} A sorted list of suggested changes.
 */
function generatePreview(directives, settings) {
  // Clear previous logs and start fresh
  clearOperationLogs();
  logOperation('INFO', `🚀 Starting preview generation for ${directives.length} directives`, {directiveCount: directives.length, settings});
  
  // Log directives for debugging
  directives.forEach((directive, index) => {
    logOperation('INFO', `📝 Directive #${index + 1}: "${directive.fragment?.substring(0, 50)}..." → "${directive.replaceWith?.substring(0, 50)}..."`, {
      directiveIndex: index,
      fragment: directive.fragment,
      replaceWith: directive.replaceWith,
      fragmentLength: directive.fragment?.length || 0,
      replaceWithLength: directive.replaceWith?.length || 0
    });
  });
  
  // Validate directives
  const validDirectives = directives.filter(d => d.fragment && d.replaceWith);
  if (validDirectives.length !== directives.length) {
    logOperation('WARN', `⚠️ ${directives.length - validDirectives.length} directives skipped due to missing fragment or replaceWith`);
  }
  logOperation('INFO', `📊 Processing ${validDirectives.length} valid directives out of ${directives.length} total`);
  
  const operationStartTime = new Date().getTime();
  const hasApiKey = checkOpenAIKey();
  logOperation('INFO', `🔑 OpenAI API key ${hasApiKey ? 'available' : 'not available'}`);
  
  try {
    // Get all document elements in one go.
    const allElements = getAllDocumentElements();
    logOperation('INFO', `📄 Retrieved ${allElements.length} document elements`, {elementCount: allElements.length});

    // PRIORITY 0: Smart Fragment Fixer - исправляет мелкие проблемы с символами
    logOperation('INFO', `🤖 Starting Smart Fragment Fixer for ${directives.length} directives`);
    const fixedDirectives = applySmartFragmentFixer(directives, allElements, hasApiKey);
    logOperation('INFO', `✅ Smart Fragment Fixer completed. Processed ${fixedDirectives.length} directives`);

    // Create a map for faster lookups
    const directiveMap = new Map();
    fixedDirectives.forEach((dir, index) => {
      directiveMap.set(index, { ...dir, normalizedFragment: normalizeText(dir.fragment) });
    });

    const elementReplacements = new Map();

    // OPTIMIZED: Single pass over elements
    allElements.forEach(elem => {
      const normElemText = normalizeText(elem.text);
      const foundInElement = [];

      directiveMap.forEach((directive, directiveIndex) => {
        if (normElemText.includes(directive.normalizedFragment)) {
          foundInElement.push({
            element: elem,
            similarity: 1.0,
            matchType: 'EXACT',
            directiveIndex: directiveIndex,
            fragment: directive.fragment,
            replaceWith: directive.replaceWith
          });
        }
      });

      if (foundInElement.length > 0) {
        elementReplacements.set(elem.originalIndex, foundInElement);
      }
    });

    const suggestions = [];
    elementReplacements.forEach((replacements, elementIndex) => {
      const firstReplacement = replacements[0];
      const element = firstReplacement.element;
      let newText = element.text;
      const fragments = [];
      const replaces = [];
      
      replacements.forEach(rep => {
        newText = newText.replace(new RegExp(escapeRegex(rep.fragment), 'g'), rep.replaceWith);
        fragments.push(rep.fragment);
        replaces.push(rep.replaceWith);
      });
      
      suggestions.push({
        type: 'EXACT',
        paraIndex: element.originalIndex,
        elementId: element.id,
        elementType: element.typeName,
        similarity: 1.0,
        oldText: element.text,
        newText: newText,
        fragment: fragments.join(' +++ '),
        replaceWith: replaces.join(' +++ '),
        directiveIndex: firstReplacement.directiveIndex, // For reference
        replacementCount: replacements.length
      });
    });

    suggestions.sort((a, b) => a.paraIndex - b.paraIndex);
    
    // PRIORITY 1: AI search for missing fragments (if API key available)
    if (hasApiKey && settings.aiThreshold > 0) {
      const foundFragments = new Set(suggestions.map(s => s.fragment.split(' +++ ')[0])); // Get first fragment from each suggestion
      const missingDirectives = fixedDirectives.filter(dir => !foundFragments.has(dir.fragment));
      
      if (missingDirectives.length > 0) {
        logOperation('INFO', `🤖 Starting AI search for ${missingDirectives.length} missing fragments`);
        
        missingDirectives.forEach((directive, index) => {
          try {
            // Use the advanced AI filtering system
            const candidates = performAdvancedAIFiltering(directive.fragment, allElements, 3);
            
            if (candidates.length > 0) {
              const aiResult = processBatchAIRequests([{
                fragment: directive.fragment,
                candidates: candidates
              }]);
              
              if (aiResult.length > 0 && aiResult[0].bestMatch) {
                const match = aiResult[0].bestMatch;
                const element = match.element;
                const newText = element.text.replace(new RegExp(escapeRegex(directive.fragment), 'g'), directive.replaceWith);
                
                suggestions.push({
                  type: 'AI',
                  paraIndex: element.originalIndex,
                  elementId: element.id,
                  elementType: element.typeName,
                  similarity: match.similarity,
                  oldText: element.text,
                  newText: newText,
                  fragment: directive.fragment,
                  replaceWith: directive.replaceWith,
                  directiveIndex: directive.originalIndex || index,
                  replacementCount: 1
                });
                
                logOperation('INFO', `✅ AI found match for "${directive.fragment.substring(0, 30)}..." in element #${element.originalIndex}`);
              }
            }
          } catch (e) {
            logOperation('WARN', `⚠️ AI search failed for fragment "${directive.fragment.substring(0, 30)}...": ${e.message}`);
          }
        });
        
        // Re-sort after adding AI suggestions
        suggestions.sort((a, b) => a.paraIndex - b.paraIndex);
        logOperation('INFO', `🤖 AI search completed. Total suggestions now: ${suggestions.length}`);
      }
    }
    
    const operationTime = (new Date().getTime() - operationStartTime) / 1000;
    
    // Log detailed statistics
    const exactSuggestions = suggestions.filter(s => s.type === 'EXACT');
    const aiSuggestions = suggestions.filter(s => s.type === 'AI');
    const totalReplacements = suggestions.reduce((sum, s) => sum + (s.replacementCount || 1), 0);
    
    logOperation('INFO', `🎉 Preview complete. Found ${suggestions.length} suggestions for ${directives.length} directives in ${operationTime.toFixed(2)}s`, {
      suggestionsCount: suggestions.length, 
      directivesCount: directives.length,
      operationTime: operationTime,
      exactSuggestions: exactSuggestions.length,
      aiSuggestions: aiSuggestions.length,
      totalReplacements: totalReplacements,
      successRate: Math.round((suggestions.length / directives.length) * 100)
    });
    
    // Add Smart Fragment Fixer statistics
    const fixerStats = {
      totalDirectives: fixedDirectives.length,
      totalFixed: fixedDirectives.filter(d => d.wasFixed).length,
      successRate: Math.round((fixedDirectives.filter(d => d.wasFixed).length / fixedDirectives.length) * 100)
    };
    
    logOperation('INFO', `📊 Smart Fragment Fixer final stats: ${fixerStats.totalFixed}/${fixerStats.totalDirectives} fixed (${fixerStats.successRate}%)`, fixerStats);
    
    // Add fixer stats to the result (pass via first suggestion)
    if (suggestions.length > 0) {
      suggestions[0].fixerStats = fixerStats;
    }

    return suggestions;
  } catch (e) {
    logOperation('ERROR', `💥💥💥 CRITICAL ERROR in generatePreview: ${e.message}`, {error: e, stack: e.stack});
    throw e;
  }
}

/**
 * Applies the approved changes to the document.
 * @param {Array<Object>} approvedSuggestions The suggestions confirmed by the user.
 * @returns {string} A summary of the operation.
 */
function applySuggestions(approvedSuggestions) {
  logOperation('INFO', "🔥 Applying approved suggestions...");
  if (!approvedSuggestions || approvedSuggestions.length === 0) return "❌ No changes selected.";
  
  const allElements = getAllDocumentElements();
  let appliedCount = 0;
  const backupEntries = [];

  const sortedSuggestions = [...approvedSuggestions].sort((a, b) => b.paraIndex - a.paraIndex);

  sortedSuggestions.forEach(suggestion => {
    try {
      const elementWrapper = allElements.find(e => e.originalIndex === suggestion.paraIndex);
      if (elementWrapper) {
        const element = elementWrapper.element;
        const currentText = getElementText(element);
        
        backupEntries.push({ paraIndex: suggestion.paraIndex, oldText: currentText, newText: suggestion.newText });
        applyTextChange(element, suggestion.newText);
        appliedCount++;
      }
    } catch (e) {
      logOperation('ERROR', `💥 ERROR applying change to element #${suggestion.paraIndex}: ${e.message}`);
    }
  });

  if (backupEntries.length > 0) {
      PropertiesService.getDocumentProperties().setProperty('LAST_RUN_BACKUP', JSON.stringify(backupEntries));
  }

  return `🎉 Applied ${appliedCount}/${sortedSuggestions.length} suggestions.`;
}

/**
 * Reverts the last set of applied changes.
 * @returns {string} A summary of the undo operation.
 */
function undoLastRun() {
  const propService = PropertiesService.getDocumentProperties();
  const backupJson = propService.getProperty('LAST_RUN_BACKUP');
  if (!backupJson) return '❌ No saved run found to undo.';

  const backupEntries = JSON.parse(backupJson);
  const allElements = getAllDocumentElements();
  let undone = 0;

  backupEntries.sort((a,b) => b.paraIndex - a.paraIndex).forEach(entry => {
    const elementWrapper = allElements.find(e => e.originalIndex === entry.paraIndex);
    if (elementWrapper) {
      applyTextChange(elementWrapper.element, entry.oldText);
      undone++;
    }
  });

  propService.deleteProperty('LAST_RUN_BACKUP');
  return `↩️ Reverted ${undone}/${backupEntries.length} changes.`;
}

/**
 * Retrieves the current progress of a running task.
 * @returns {Object|null} The progress status.
 */
function getRitualProgress() {
  const prop = PropertiesService.getDocumentProperties().getProperty('RITUAL_PROGRESS');
  return prop ? JSON.parse(prop) : null;
}

/**
 * Получает детальную информацию о статусе API ключа, включая миграцию
 * @returns {Object} Информация о статусе API ключа
 */
function getApiKeyStatus() {
  try {
    const documentKey = PropertiesService.getDocumentProperties().getProperty('OPENAI_API_KEY');
    const scriptKey = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    
    let status = {
      hasKey: false,
      source: 'none',
      migrated: false,
      message: ''
    };
    
    if (documentKey && scriptKey) {
      status.hasKey = true;
      status.source = 'both';
      status.message = 'API ключ найден в обеих локациях';
    } else if (documentKey) {
      status.hasKey = true;
      status.source = 'document';
      status.message = 'API ключ найден в Document Properties';
    } else if (scriptKey) {
      status.hasKey = true;
      status.source = 'script';
      status.migrated = true;
      status.message = 'API ключ найден в Script Properties и автоматически мигрирован';
      // Выполняем миграцию
      PropertiesService.getDocumentProperties().setProperty('OPENAI_API_KEY', scriptKey);
    } else {
      status.hasKey = false;
      status.source = 'none';
      status.message = 'API ключ не найден';
    }
    
    logOperation('INFO', `🔑 API key status: ${status.message}`, status);
    return status;
  } catch (e) {
    logOperation('ERROR', `🔑 Error getting API key status: ${e.message}`);
    return { hasKey: false, source: 'error', migrated: false, message: 'Ошибка проверки API ключа' };
  }
}
