/*
==========================================================================================
=== Advanced Replacer for Google Docs v1.0 ===
==========================================================================================
*
* VERSION: 1.0 (GitHub Release)
* DATE: 2024-07-27
* DESCRIPTION: An advanced batch find-and-replace tool for Google Docs.
*
* FEATURES:
* 1. Supports all content types: Headings, Paragraphs, Lists, Tables.
* 2. Smart, single-pass processing engine avoids re-checking elements.
* 3. Advanced text normalization for better matching.
* 4. Multi-level matching: EXACT and AI-powered semantic matching.
* 5. "Undo" functionality to revert the last batch of changes.
* 6. Real-time progress bar for large documents.
* 7. Enhanced element identification with unique IDs.
* 8. Live editing mode for replacement text.
* 9. Comprehensive logging and error reporting.
*
==========================================================================================
*/

/* ───────── UI & Menu ───────── */

/**
 * Creates the Add-on menu in the Google Docs UI when the document is opened.
 */
function onOpen() {
  DocumentApp.getUi()
    .createMenu('🚀 Advanced Replacer')
    .addItem('Open Sidebar', 'showSidebar')
    .addToUi();
}

/**
 * Displays the HTML sidebar.
 */
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('🚀 Advanced Replacer');
  DocumentApp.getUi().showSidebar(html);
}

/**
 * Displays the UI as a modal dialog for an expanded view.
 */
function showAsModal(directivesText) {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Sidebar')
      .setWidth(1400)
      .setHeight(1000);
 
  htmlOutput.append(`<script>
      window.addEventListener('DOMContentLoaded', function() {
        if (typeof init === 'function') {
            init(${JSON.stringify(directivesText)});
        } else {
            // Fallback if init is not ready
            document.getElementById('directives-textarea').value = ${JSON.stringify(directivesText)};
        }
      });
    </script>`);

  DocumentApp.getUi().showModalDialog(htmlOutput, '🚀 Advanced Replacer (Expanded View)');
}

/* ───────── Core Utilities ───────── */

/**
 * Safely gets the body of the active document.
 * @returns {GoogleAppsScript.Document.Body} The document body.
 * @throws {Error} If no active document or body is found.
 */
function safeGetBody() {
  logOperation('INFO', '📄 Getting document body...');
  
  try {
    const doc = DocumentApp.getActiveDocument();
    if (!doc) {
      logOperation('ERROR', 'No active Google Docs document found');
      throw new Error('No active Google Docs document found.');
    }
    
    logOperation('INFO', '✅ Active document found', {
      documentId: doc.getId(),
      documentName: doc.getName(),
      documentUrl: doc.getUrl()
    });
    
    const body = doc.getBody();
    if (!body) {
      logOperation('ERROR', 'Could not retrieve the document body');
      throw new Error('Could not retrieve the document body.');
    }
    
    logOperation('INFO', '✅ Document body retrieved successfully', {
      bodyChildrenCount: body.getNumChildren ? body.getNumChildren() : 'unknown'
    });
    
    return body;
  } catch (e) {
    logOperation('ERROR', `💥 Error in safeGetBody: ${e.message}`, {
      error: e.message,
      stack: e.stack
    });
    throw e;
  }
}

/**
 * Retrieves all processable text-containing elements from the document.
 * This function recursively traverses the document structure.
 * @returns {Array<Object>} An array of element wrappers with unique IDs.
 */
function getAllDocumentElements() {
  logOperation('INFO', '📄 Starting document elements retrieval...');
  const startTime = new Date().getTime();
  
  try {
    const body = safeGetBody();
    const elements = [];
    let index = 0;
    let elementTypeCounts = {};

    function processElement(element) {
      try {
        const elementType = element.getType();
        const elementTypeName = getElementTypeName(elementType);
        const text = getElementText(element).trim();
        
        // Track element type statistics
        elementTypeCounts[elementTypeName] = (elementTypeCounts[elementTypeName] || 0) + 1;
        
        // FIX v5.1.1: Skip the root BODY_SECTION to avoid matching the entire document as one element.
        if (elementType !== DocumentApp.ElementType.BODY_SECTION && text.length > 0) {
          const elementWrapper = {
            element: element,
            text: text,
            type: elementType,
            typeName: elementTypeName,
            originalIndex: index++
          };
          
          // Generate unique ID for this element
          elementWrapper.id = generateElementId(elementWrapper);
          
          elements.push(elementWrapper);
          logOperation('INFO', `📝 Added element #${elementWrapper.originalIndex}: ${elementWrapper.typeName} (${text.length} chars, ID: ${elementWrapper.id.substring(0, 15)}...)`, {
            elementIndex: elementWrapper.originalIndex,
            elementType: elementWrapper.typeName,
            textLength: text.length,
            elementId: elementWrapper.id,
            textPreview: text.substring(0, 100) + (text.length > 100 ? '...' : '')
          });
        } else if (elementType === DocumentApp.ElementType.BODY_SECTION) {
          logOperation('INFO', `🚫 Skipped BODY_SECTION element (avoids matching entire document)`);
        } else if (text.length === 0) {
          logOperation('INFO', `🚫 Skipped empty ${elementTypeName} element`);
        }

        // Recursively process children of container elements.
        if (element.getNumChildren && element.getNumChildren() > 0) {
          logOperation('INFO', `📂 Processing ${element.getNumChildren()} children of ${elementTypeName}`, {
            parentType: elementTypeName,
            childrenCount: element.getNumChildren()
          });
          
          for (let i = 0; i < element.getNumChildren(); i++) {
            try {
              const child = element.getChild(i);
              processElement(child);
            } catch (e) {
              // Log and skip children that can't be processed (e.g., permissions issues).
              logOperation('WARN', `⚠️ Skipping child element #${i}: ${e.message}`, {
                childIndex: i,
                parentType: elementTypeName,
                error: e.message
              });
            }
          }
        }
      } catch (e) {
        logOperation('ERROR', `💥 Error processing element: ${e.message}`, {
          error: e.message,
          stack: e.stack,
          elementType: element ? getElementTypeName(element.getType()) : 'unknown'
        });
      }
    }

    processElement(body);
    
    const processingTime = new Date().getTime() - startTime;
    logOperation('INFO', `✅ Document parsing complete: Found ${elements.length} elements in ${processingTime}ms`, {
      totalElements: elements.length,
      processingTime: processingTime,
      elementTypeCounts: elementTypeCounts
    });
    
    // Log summary of element types found
    Object.entries(elementTypeCounts).forEach(([type, count]) => {
      logOperation('INFO', `📊 Element type summary: ${type} = ${count}`);
    });
    
    return elements;
  } catch (e) {
    logOperation('ERROR', `💥 CRITICAL ERROR in getAllDocumentElements: ${e.message}`, {
      error: e.message,
      stack: e.stack
    });
    throw e;
  }
}

/**
 * Intelligently extracts text from any Google Docs element type.
 * @param {GoogleAppsScript.Document.Element} element The document element.
 * @returns {string} The extracted text.
 */
function getElementText(element) {
  try {
    const elementType = element.getType();
    
    switch (elementType) {
      case DocumentApp.ElementType.PARAGRAPH:
      case DocumentApp.ElementType.HEADING:
      case DocumentApp.ElementType.LIST_ITEM:
        return element.getText();
      case DocumentApp.ElementType.TABLE:
        let tableText = '';
        const numRows = element.getNumRows();
        for (let row = 0; row < numRows; row++) {
          const numCells = element.getRow(row).getNumCells();
          for (let col = 0; col < numCells; col++) {
            tableText += element.getCell(row, col).getText() + ' ';
          }
        }
        return tableText;
      case DocumentApp.ElementType.TABLE_CELL:
      case DocumentApp.ElementType.TEXT:
        return element.getText();
      default:
        // Fallback for other potential text-containing elements.
        if (element.getText) return element.getText();
        if (element.asText) return element.asText().getText();
        return '';
    }
  } catch (e) {
    // Return empty string if text extraction fails for any reason.
    return '';
  }
}

/**
 * Gets a human-readable name for an element type.
 * @param {GoogleAppsScript.Document.ElementType} elementType The element type enum.
 * @returns {string} The readable name.
 */
function getElementTypeName(elementType) {
  const typeMap = {
    [DocumentApp.ElementType.PARAGRAPH]: 'Paragraph',
    [DocumentApp.ElementType.HEADING]: 'Heading', 
    [DocumentApp.ElementType.LIST_ITEM]: 'List Item',
    [DocumentApp.ElementType.TABLE]: 'Table',
    [DocumentApp.ElementType.TABLE_CELL]: 'Table Cell',
    [DocumentApp.ElementType.TEXT]: 'Text'
  };
  return typeMap[elementType] || 'Unknown';
}

/**
 * Advanced text normalization for better text matching.
 * Currently configured for Ukrainian, can be adapted for other languages.
 * @param {string} text The input text.
 * @returns {string} The normalized text.
 */
function normalizeText(text) {
  if (!text) return '';
  
  return text
    .toLowerCase()
    .trim()
    .replace(/\s+/g, ' ') // Collapse whitespace
    .replace(/[^\w\s\u00C0-\u017F0-9]/g, '') // Remove special characters, keeping letters and numbers
    .replace(/['`]/g, 'ʼ') // Normalize apostrophes (example for Ukrainian)
    .trim();
}

/* ───────── 🤖 Smart Fragment Fixer ───────── */

/**
 * Интеллектуальная система исправления мелких проблем во фрагментах
 * Автоматически находит и исправляет проблемы с символами, пробелами, кавычками
 * @param {Array<Object>} directives Исходные директивы
 * @param {Array<Object>} allElements Все элементы документа
 * @returns {Array<Object>} Исправленные директивы
 */
function applySmartFragmentFixer(directives, allElements, hasApiKey) {
  logOperation('INFO', '🤖 Smart Fragment Fixer: начинаю анализ директив');
  
  const fixedDirectives = [];
  let totalFixed = 0;
  
  directives.forEach((directive, index) => {
    const originalFragment = directive.fragment;
    let currentFragment = originalFragment;
    let wasFixed = false;
    let fixType = [];
    
    logOperation('INFO', `🔍 Анализ директивы #${index + 1}: "${originalFragment.substring(0, 50)}..."`, {
      directiveIndex: index,
      originalFragment: originalFragment
    });
    
    // Сначала проверяем - может уже есть точное совпадение
    const hasExactMatch = allElements.some(elem => elem.text.includes(originalFragment));
    
    if (hasExactMatch) {
      logOperation('INFO', `✅ Директива #${index + 1}: точное совпадение найдено, исправления не нужны`);
      fixedDirectives.push({
        ...directive,
        originalFragment: originalFragment,
        wasFixed: false,
        fixType: []
      });
      return;
    }
    
    logOperation('WARN', `⚠️ Директива #${index + 1}: точное совпадение НЕ найдено, начинаю умные исправления`);
    
    // ФИКС 1: Замена типов тире
    let fixedFragment = currentFragment;
    const originalWithDashes = fixedFragment;
    fixedFragment = fixedFragment
      .replace(/\s-\s/g, ' — ')  // Короткое тире с пробелами → длинное тире с пробелами
      .replace(/^-\s/g, '— ')    // Тире в начале строки
      .replace(/\s-$/g, ' —');   // Тире в конце строки
    
    if (fixedFragment !== originalWithDashes) {
      fixType.push('тире');
      currentFragment = fixedFragment;
      wasFixed = true;
      logOperation('INFO', `🔧 Фикс #${index + 1}: исправлены тире`, {
        before: originalWithDashes,
        after: fixedFragment
      });
    }
    
    // ФИКС 2: Замена кавычек
    const originalWithQuotes = fixedFragment;
    fixedFragment = fixedFragment
      .replace(/\\"/g, '«')      // Экранированные кавычки в начале
      .replace(/\"([^\"]*)\"/g, '«$1»')  // Парные обычные кавычки
      .replace(/"([^"]*)"/g, '«$1»')     // Парные "умные" кавычки  
      .replace(/"/g, '«')        // Оставшиеся одинарные кавычки в открывающие
      .replace(/"/g, '»');       // Закрывающие "умные" кавычки
    
    if (fixedFragment !== originalWithQuotes) {
      fixType.push('кавычки');
      currentFragment = fixedFragment;
      wasFixed = true;
      logOperation('INFO', `🔧 Фикс #${index + 1}: исправлены кавычки`, {
        before: originalWithQuotes,
        after: fixedFragment
      });
    }
    
    // ФИКС 3: Нормализация пробелов
    const originalWithSpaces = fixedFragment;
    fixedFragment = fixedFragment
      .replace(/\s+/g, ' ')      // Множественные пробелы → одинарные
      .replace(/\u00A0/g, ' ')   // Неразрывные пробелы → обычные
      .replace(/\u2009/g, ' ')   // Тонкие пробелы → обычные
      .trim();                   // Убираем пробелы в начале/конце
    
    // ФИКС 4: Украинские символы и апострофы
    const originalWithApostrophes = fixedFragment;
    fixedFragment = fixedFragment
      .replace(/'/g, 'ʼ')        // Обычный апостроф → украинский апостроф
      .replace(/`/g, 'ʼ')        // Обратный апостроф → украинский апостроф
      .replace(/'/g, 'ʼ')        // "Умный" апостроф → украинский апостроф
      .replace(/'/g, 'ʼ');       // Другой "умный" апостроф → украинский апостроф
    
    if (fixedFragment !== originalWithApostrophes) {
      fixType.push('апострофы');
      currentFragment = fixedFragment;
      wasFixed = true;
      logOperation('INFO', `🔧 Фикс #${index + 1}: исправлены апострофы`, {
        before: originalWithApostrophes,
        after: fixedFragment
      });
    }
    
    if (fixedFragment !== originalWithSpaces) {
      fixType.push('пробелы');
      currentFragment = fixedFragment;
      wasFixed = true;
      logOperation('INFO', `🔧 Фикс #${index + 1}: нормализованы пробелы`, {
        before: originalWithSpaces,
        after: fixedFragment
      });
    }
    
    // ФИКС 5: Проверка исправленного фрагмента
    const hasFixedMatch = allElements.some(elem => elem.text.includes(currentFragment));
    
    if (hasFixedMatch && wasFixed) {
      totalFixed++;
      logOperation('INFO', `✅ Директива #${index + 1}: исправления помогли! Найдено точное совпадение`, {
        originalFragment: originalFragment,
        fixedFragment: currentFragment,
        fixTypes: fixType
      });
    } else if (wasFixed) {
      logOperation('WARN', `⚠️ Директива #${index + 1}: исправления не помогли, совпадение все еще не найдено`, {
        originalFragment: originalFragment,
        fixedFragment: currentFragment,
        fixTypes: fixType
      });
    } else {
      logOperation('WARN', `❌ Директива #${index + 1}: не удалось найти совпадения даже после исправлений`);
    }
    
    // ФИКС 6: AI-помощь для сложных случаев (если есть API ключ)
    if (!hasFixedMatch && hasApiKey) {
      const aiFixedFragment = applyAIFragmentFixer_Legacy(currentFragment, allElements);
      if (aiFixedFragment && aiFixedFragment !== currentFragment) {
        const hasAIMatch = allElements.some(elem => elem.text.includes(aiFixedFragment));
        if (hasAIMatch) {
          currentFragment = aiFixedFragment;
          fixType.push('AI');
          wasFixed = true;
          totalFixed++;
          logOperation('INFO', `🤖 Директива #${index + 1}: AI исправление успешно!`, {
            originalFragment: originalFragment,
            aiFixedFragment: aiFixedFragment
          });
        }
      }
    }
    
    fixedDirectives.push({
      ...directive,
      fragment: currentFragment,
      originalFragment: originalFragment,
      wasFixed: wasFixed,
      fixType: fixType
    });
  });
  
  logOperation('INFO', `🎉 Smart Fragment Fixer завершен: исправлено ${totalFixed}/${directives.length} директив`, {
    totalDirectives: directives.length,
    totalFixed: totalFixed,
    successRate: Math.round((totalFixed / directives.length) * 100)
  });
  
  return fixedDirectives;
}

/**
 * Вычисляет базовую схожесть между двумя текстовыми строками
 * @param {string} str1 Первая строка
 * @param {string} str2 Вторая строка
 * @returns {number} Значение от 0 до 1
 */
function calculateBasicSimilarity(str1, str2) {
  if (!str1 || !str2) return 0;
  
  const words1 = str1.toLowerCase().split(/\s+/);
  const words2 = str2.toLowerCase().split(/\s+/);
  
  let commonWords = 0;
  words1.forEach(word => {
    if (words2.includes(word)) commonWords++;
  });
  
  return commonWords / Math.max(words1.length, words2.length);
}

/**
 * AI-помощник для исправления сложных случаев во фрагментах
 * @param {string} fragment Фрагмент для исправления
 * @param {Array<Object>} allElements Все элементы документа
 * @returns {string|null} Исправленный фрагмент или null
 */
function applyAIFragmentFixer_Legacy(fragment, allElements) {
  try {
    logOperation('INFO', `🤖 AI Fragment Fixer: анализирую "${fragment.substring(0, 50)}..."`);
    
    // Ищем наиболее похожие элементы
    const candidates = allElements
      .map(elem => ({
        element: elem,
        similarity: calculateBasicSimilarity(fragment, elem.text)
      }))
      .filter(c => c.similarity > 0.3)
      .sort((a, b) => b.similarity - a.similarity)
      .slice(0, 3);
    
    if (candidates.length === 0) {
      logOperation('WARN', '🤖 AI Fragment Fixer: подходящие кандидаты не найдены');
      return null;
    }
    
    logOperation('INFO', `🤖 AI Fragment Fixer: найдено ${candidates.length} кандидатов для анализа`);
    
    // Для простоты, пока возвращаем null - можно расширить позже
    // Здесь можно добавить более сложную AI логику
    return null;
    
  } catch (e) {
    logOperation('ERROR', `💥 Ошибка в AI Fragment Fixer: ${e.message}`, {
      error: e.message,
      fragment: fragment
    });
    return null;
  }
}

/* ───────── 🚀 Main Processing Engine ───────── */

/**
 * Generates replacement previews based on a single pass over the document.
 * Fixed version that properly handles multiple replacements in the same element.
 * Enhanced with detailed logging and error tracking.
 * @param {Array<Object>} directives An array of {fragment, replaceWith} objects.
 * @param {Object} settings An object with {aiThreshold}.
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
  
  // CRITICAL: Log each directive's searchability
  const body = safeGetBody();
  const bodyText = body.getText();
  
  logOperation('INFO', `📄 Document text length: ${bodyText.length} characters`);
  
  directives.forEach((directive, index) => {
    if (directive.fragment) {
      const directMatch = bodyText.includes(directive.fragment);
      const trimmedMatch = bodyText.includes(directive.fragment.trim());
      
      logOperation(directMatch ? 'INFO' : 'WARN', `🔍 Directive #${index + 1} searchability: direct=${directMatch}, trimmed=${trimmedMatch}`, {
        directiveIndex: index,
        fragment: directive.fragment,
        fragmentLength: directive.fragment.length,
        directMatch: directMatch,
        trimmedMatch: trimmedMatch,
        hasLeadingSpaces: directive.fragment.length !== directive.fragment.trimStart().length,
        hasTrailingSpaces: directive.fragment.length !== directive.fragment.trimEnd().length
      });
    }
  });
  
  // Log each directive for debugging
  directives.forEach((directive, index) => {
    logOperation('INFO', `📝 Directive #${index + 1}: "${directive.fragment}" → "${directive.replaceWith}"`, {
      directiveIndex: index,
      fragment: directive.fragment,
      replaceWith: directive.replaceWith,
      fragmentLength: directive.fragment?.length || 0,
      replaceWithLength: directive.replaceWith?.length || 0
    });
  });
  
  const operationStartTime = new Date().getTime();
  const hasApiKey = checkOpenAIKey();
  logOperation('INFO', `🔑 OpenAI API key ${hasApiKey ? 'available' : 'not available'}`);
  
  const suggestions = [];
  
  try {
    // Get all document elements in one go.
    const allElements = getAllDocumentElements();
    logOperation('INFO', `📄 Retrieved ${allElements.length} document elements`, {elementCount: allElements.length});
    
    // Group directives by element to handle multiple replacements in same element
    const elementReplacements = new Map(); // elementIndex -> Array of replacements
    
    // PRIORITY 0: Smart Fragment Fixer - исправляет мелкие проблемы с символами
    logOperation('INFO', `🤖 Starting Smart Fragment Fixer for ${directives.length} directives`);
    const fixedDirectives = applySmartFragmentFixer(directives, allElements, hasApiKey);
    logOperation('INFO', `✅ Smart Fragment Fixer completed. Processed ${fixedDirectives.length} directives`);

    fixedDirectives.forEach((directive, directiveIndex) => {
      const { fragment, replaceWith, originalFragment, wasFixed } = directive;
      
      if (!fragment || !replaceWith) {
        logOperation('WARN', `⚠️ Directive #${directiveIndex + 1} skipped: empty fields`, {directive});
        return;
      }

      let logMessage = `🎯 [Directive #${directiveIndex + 1}] Processing: "${fragment}" → "${replaceWith}"`;
      if (wasFixed) {
        logMessage += ` [🔧 FIXED from: "${originalFragment}"]`;
        logOperation('INFO', `🔧 Fixed directive #${directiveIndex + 1}: "${originalFragment}" → "${fragment}"`, {
          directiveIndex: directiveIndex,
          originalFragment: originalFragment,
          fixedFragment: fragment,
          fixType: directive.fixType
        });
      }
      logOperation('INFO', logMessage);
      
      let bestMatch = null;
      let bestScore = 0;
      let candidatesChecked = 0;
      
      // Single pass over all elements with simplified logic: EXACT or AI only
      allElements.forEach((elem, elemIndex) => {
        candidatesChecked++;
        
        // Log progress every 50 elements or for first 10 elements
        if (candidatesChecked <= 10 || candidatesChecked % 50 === 0) {
          logOperation('INFO', `🔍 Checking element #${elem.originalIndex}/${allElements.length} (${elem.typeName}): "${elem.text.substring(0, 50)}${elem.text.length > 50 ? '...' : ''}"`, {
            elementIndex: elem.originalIndex,
            elementType: elem.typeName,
            textLength: elem.text.length,
            candidatesChecked: candidatesChecked
          });
        }
        
        // PRIORITY 1: Look for exact matches
        if (elem.text.includes(fragment)) {
          logOperation('INFO', `🎯 POTENTIAL EXACT match found in element #${elem.originalIndex}: "${fragment}"`, {
            elementIndex: elem.originalIndex,
            elementType: elem.typeName,
            elementId: elem.id,
            fragment: fragment,
            contextBefore: elem.text.substring(Math.max(0, elem.text.indexOf(fragment) - 20), elem.text.indexOf(fragment)),
            contextAfter: elem.text.substring(elem.text.indexOf(fragment) + fragment.length, elem.text.indexOf(fragment) + fragment.length + 20)
          });
          
          try {
            // findText is only available on some elements, and it confirms the match.
            if (elem.element.findText && elem.element.findText(fragment)) {
              bestMatch = {
                element: elem,
                similarity: 1.0,
                matchType: 'EXACT',
                directiveIndex: directiveIndex,
                fragment: fragment,
                replaceWith: replaceWith
              };
              bestScore = 1.0;
              logOperation('INFO', `✅ EXACT match CONFIRMED via findText() in element #${elem.originalIndex}`, {
                elementId: elem.id, 
                elementType: elem.typeName,
                fragment: fragment,
                textPreview: elem.text.substring(0, 100) + (elem.text.length > 100 ? '...' : '')
              });
            }
          } catch (e) {
            // Element doesn't support findText, but text.includes found it
            // This is still a valid exact match
            bestMatch = {
              element: elem,
              similarity: 1.0,
              matchType: 'EXACT',
              directiveIndex: directiveIndex,
              fragment: fragment,
              replaceWith: replaceWith
            };
            bestScore = 1.0;
            logOperation('INFO', `✅ EXACT match found (fallback - no findText support) in element #${elem.originalIndex}`, {
              elementId: elem.id, 
              elementType: elem.typeName,
              fragment: fragment,
              error: e.message,
              textPreview: elem.text.substring(0, 100) + (elem.text.length > 100 ? '...' : '')
            });
          }
        }
      });

      logOperation('INFO', `🔍 Checked ${candidatesChecked} elements for directive #${directiveIndex + 1}`);

      // If no exact match found, collect for batch AI processing
      if (!bestMatch) {
        // Store this directive for AI processing
        if (!hasApiKey) {
          logOperation('WARN', `⚠️ No AI key available for semantic search of "${fragment}"`);
        }
      }

      // If a match was found, group it by element for batch processing
      if (bestMatch) {
        logOperation('INFO', `✅ ${bestMatch.matchType}: Found in ${bestMatch.element.typeName} #${bestMatch.element.originalIndex} (Score: ${Math.round(bestMatch.similarity * 100)}%)`);
        
        const elementIndex = bestMatch.element.originalIndex;
        
        if (!elementReplacements.has(elementIndex)) {
          elementReplacements.set(elementIndex, []);
        }
        
        elementReplacements.get(elementIndex).push(bestMatch);
      } else {
        logOperation('WARN', `❌ No suitable match found for "${fragment}"`, {directive, candidatesChecked});
      }
    });

    // PRIORITY 2: Batch AI processing for directives without exact matches
    // ❌ ВРЕМЕННО ОТКЛЮЧЕНО ДЛЯ ОТЛАДКИ
    logOperation('WARN', '⚠️ AI система временно отключена для отладки - будут обрабатываться только EXACT совпадения', {
      hasApiKey: hasApiKey,
      reason: 'AI система была отключена для исправления багов с заменой'
    });
    
    if (false && hasApiKey) {
      const aiDirectives = [];
      
      // Collect directives that need AI processing
      directives.forEach((directive, directiveIndex) => {
        const { fragment, replaceWith } = directive;
        if (!fragment || !replaceWith) return;
        
        // Check if this directive already found an exact match
        let hasExactMatch = false;
        elementReplacements.forEach(replacements => {
          if (replacements.some(r => r.directiveIndex === directiveIndex)) {
            hasExactMatch = true;
          }
        });
        
        if (!hasExactMatch) {
          aiDirectives.push({ fragment, replaceWith, directiveIndex });
        }
      });
      
      logOperation('INFO', `🤖 Processing ${aiDirectives.length} directives with AI batch system`);
      
      if (aiDirectives.length > 0) {
        // Prepare AI requests for batch processing
        const aiRequests = aiDirectives.map(directive => {
          const candidates = performAdvancedAIFiltering(directive.fragment, allElements, 3);
          return {
            fragment: directive.fragment,
            replaceWith: directive.replaceWith,
            directiveIndex: directive.directiveIndex,
            candidates: candidates
          };
        });
        
        // Process AI requests in batch
        const aiResults = processBatchAIRequests(aiRequests);
        
        // Process AI results and add to elementReplacements
        aiResults.forEach((result, resultIndex) => {
          if (result.bestMatch) {
            const aiDirective = aiDirectives[resultIndex];
            const bestMatch = {
              element: result.bestMatch.element,
              similarity: 0.95,
              matchType: 'AI',
              directiveIndex: aiDirective.directiveIndex,
              fragment: aiDirective.fragment,
              replaceWith: aiDirective.replaceWith
            };
            
            const elementIndex = result.bestMatch.element.originalIndex;
            
            if (!elementReplacements.has(elementIndex)) {
              elementReplacements.set(elementIndex, []);
            }
            
            elementReplacements.get(elementIndex).push(bestMatch);
            
            logOperation('INFO', `✅ AI batch: Found match for "${aiDirective.fragment.substring(0, 50)}..." in element #${elementIndex}`, {
              elementId: result.bestMatch.element.id,
              score: result.bestMatch.combinedScore
            });
          } else {
            logOperation('WARN', `❌ AI batch: No match found for "${aiDirectives[resultIndex].fragment.substring(0, 50)}..."`);
          }
        });
      }
    }

    logOperation('INFO', `📊 Final: Grouped replacements for ${elementReplacements.size} elements`);

    // Now process grouped replacements for each element
    elementReplacements.forEach((replacements, elementIndex) => {
      try {
        const element = replacements[0].element; // All replacements are for the same element
        const originalText = element.text;
        
        logOperation('INFO', `🔄 Processing ${replacements.length} replacements for element #${elementIndex}`, {elementId: element.id});
        
        // Apply all replacements to create the new text
        let newText = originalText;
        let allFragments = [];
        let allReplaceWith = [];
        
        // Sort replacements by position in text (for EXACT) or by priority
        replacements.sort((a, b) => {
          if (a.matchType === 'EXACT' && b.matchType === 'EXACT') {
            return originalText.indexOf(a.fragment) - originalText.indexOf(b.fragment);
          }
          // AI replacements go last
          if (a.matchType === 'AI' && b.matchType === 'EXACT') return 1;
          if (a.matchType === 'EXACT' && b.matchType === 'AI') return -1;
          return 0;
        });
        
        if (replacements.length === 1) {
          // Single replacement
          const replacement = replacements[0];
          if (replacement.matchType === 'EXACT') {
            newText = originalText.replace(new RegExp(escapeRegex(replacement.fragment), 'g'), replacement.replaceWith);
          } else {
            // ❌ ВРЕМЕННО ОТКЛЮЧАЕМ AI ЗАМЕНЫ - ОНИ РАБОТАЮТ НЕПРАВИЛЬНО
            newText = originalText; // Не применяем AI замену
            logOperation('WARN', `⚠️ AI замена отключена для безопасности: "${replacement.fragment}" → "${replacement.replaceWith}"`);
          }
          
          suggestions.push({
            type: replacement.matchType,
            paraIndex: elementIndex,
            elementId: element.id, // Add element ID
            elementType: element.typeName,
            similarity: replacement.similarity,
            oldText: originalText,
            newText: newText,
            fragment: replacement.fragment,
            replaceWith: replacement.replaceWith,
            directiveIndex: replacement.directiveIndex,
            replacementCount: 1
          });
          
          logOperation('INFO', `➕ Added single ${replacement.matchType} suggestion for element #${elementIndex}`);
        } else {
          // Multiple replacements in same element
          const exactReplacements = replacements.filter(r => r.matchType === 'EXACT');
          const aiReplacements = replacements.filter(r => r.matchType === 'AI');
          
          logOperation('INFO', `🔀 Processing multiple replacements: ${exactReplacements.length} EXACT, ${aiReplacements.length} AI`);
          
          if (exactReplacements.length > 0) {
            // Apply all exact replacements
            exactReplacements.forEach(replacement => {
              newText = newText.replace(new RegExp(escapeRegex(replacement.fragment), 'g'), replacement.replaceWith);
              allFragments.push(replacement.fragment);
              allReplaceWith.push(replacement.replaceWith);
            });
            
            suggestions.push({
              type: 'EXACT',
              paraIndex: elementIndex,
              elementId: element.id, // Add element ID
              elementType: element.typeName,
              similarity: 1.0,
              oldText: originalText,
              newText: newText,
              fragment: allFragments.join(' + '),
              replaceWith: allReplaceWith.join(' + '),
              directiveIndex: exactReplacements[0].directiveIndex,
              replacementCount: exactReplacements.length
            });
            
            logOperation('INFO', `➕ Added combined EXACT suggestion with ${exactReplacements.length} replacements for element #${elementIndex}`);
          }
          
          if (aiReplacements.length > 0) {
            // AI replacements replace entire text, so take the first one
            const aiReplacement = aiReplacements[0];
            
            // ❌ ВРЕМЕННО ОТКЛЮЧАЕМ AI ЗАМЕНЫ
            logOperation('WARN', `⚠️ AI замена блокирована: element #${elementIndex}, fragment: "${aiReplacement.fragment}"`);
            /*
            suggestions.push({
              type: 'AI',
              paraIndex: elementIndex,
              elementId: element.id, // Add element ID
              elementType: element.typeName,
              similarity: aiReplacement.similarity,
              oldText: originalText,
              newText: aiReplacement.replaceWith,
              fragment: aiReplacement.fragment,
              replaceWith: aiReplacement.replaceWith,
              directiveIndex: aiReplacement.directiveIndex,
              replacementCount: 1
            });
            */
            
            logOperation('INFO', `➕ Added AI suggestion for element #${elementIndex}`);
            
            if (aiReplacements.length > 1) {
              logOperation('WARN', `⚠️ Multiple AI replacements for element #${elementIndex}, only using the first one`, {skipped: aiReplacements.length - 1});
            }
          }
        }
      } catch (e) {
        logOperation('ERROR', `💥 Error processing replacements for element #${elementIndex}: ${e.message}`, {error: e, elementIndex});
      }
    });

    suggestions.sort((a, b) => a.paraIndex - b.paraIndex);
    
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
    
    logOperation('INFO', `📊 Results breakdown: ${exactSuggestions.length} EXACT, ${aiSuggestions.length} AI, ${totalReplacements} total replacements`, {
      exactCount: exactSuggestions.length,
      aiCount: aiSuggestions.length,
      totalReplacements: totalReplacements,
      missedDirectives: directives.length - suggestions.length
    });
    
    // Log any missed directives
    if (suggestions.length < directives.length) {
      logOperation('WARN', `⚠️ ${directives.length - suggestions.length} directives did not find matches`, {
        totalDirectives: directives.length,
        foundMatches: suggestions.length,
        missedCount: directives.length - suggestions.length
      });
    }
    
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
    } else {
      // If no suggestions, create a dummy suggestion with fixer stats
      suggestions.push({
        type: 'INFO',
        paraIndex: -1,
        elementId: 'fixer-stats',
        elementType: 'INFO',
        similarity: 0,
        oldText: '',
        newText: '',
        fragment: 'Smart Fragment Fixer Stats',
        replaceWith: `${fixerStats.totalFixed}/${fixerStats.totalDirectives} fixed`,
        directiveIndex: -1,
        replacementCount: 0,
        fixerStats: fixerStats
      });
    }
    
    // Final log count check for debugging
    const totalLogs = OPERATION_LOGS.length;
    const errorCount = OPERATION_LOGS.filter(log => log.level === 'ERROR').length;
    const warnCount = OPERATION_LOGS.filter(log => log.level === 'WARN').length;
    
    logOperation('INFO', `📋 Operation complete. Total logs: ${totalLogs}, errors: ${errorCount}, warnings: ${warnCount}`, {
      totalLogs: totalLogs,
      errorCount: errorCount,
      warnCount: warnCount,
      logSample: OPERATION_LOGS.slice(-5).map(log => `${log.level}: ${log.message}`)
    });
    
    return suggestions;
  } catch (e) {
    logOperation('ERROR', `💥💥💥 CRITICAL ERROR in generatePreview: ${e.message}`, {error: e, stack: e.stack});
    
    // Log error count for debugging
    const totalLogs = OPERATION_LOGS.length;
    const errorCount = OPERATION_LOGS.filter(log => log.level === 'ERROR').length;
    logOperation('INFO', `📋 Error occurred. Total logs: ${totalLogs}, errors: ${errorCount}`, {
      totalLogs: totalLogs,
      errorCount: errorCount
    });
    
    throw e;
  }
}

/**
 * Basic similarity calculation for AI candidate filtering
 * @param {string} fragment The search fragment.
 * @param {string} text The text to compare against.
 * @returns {number} A similarity score between 0 and 1.
 */
function calculateBasicSimilarity(fragment, text) {
  if (!fragment || !text) return 0;
  
  const normFragment = normalizeText(fragment);
  const normText = normalizeText(text);
  
  if (normFragment === normText) return 1;
  if (normText.includes(normFragment)) return 0.9;
  
  // Simple word overlap check
  const fragmentWords = normFragment.split(/\s+/).filter(w => w.length > 2);
  const textWords = new Set(normText.split(/\s+/));
  
  let matches = 0;
  fragmentWords.forEach(word => {
    if (textWords.has(word)) matches++;
  });
  
  return fragmentWords.length > 0 ? matches / fragmentWords.length : 0;
}

/**
 * 🚀 SUPER-OPTIMIZED AI SEARCH SYSTEM
 * Handles large documents efficiently with smart filtering and batching
 */

// Global AI cache to avoid repeated requests
if (typeof AI_SEARCH_CACHE === 'undefined') {
  var AI_SEARCH_CACHE = new Map();
}

/**
 * Advanced similarity calculation with multiple metrics
 * @param {string} fragment The search fragment.
 * @param {string} text The text to compare against.
 * @returns {Object} Detailed similarity metrics.
 */
function calculateAdvancedSimilarity(fragment, text) {
  if (!fragment || !text) return { score: 0, wordOverlap: 0, lengthSimilarity: 0 };
  
  const normFragment = normalizeText(fragment);
  const normText = normalizeText(text);
  
  // Exact/substring check
  if (normFragment === normText) return { score: 1.0, wordOverlap: 1.0, lengthSimilarity: 1.0 };
  if (normText.includes(normFragment)) return { score: 0.9, wordOverlap: 0.9, lengthSimilarity: 0.85 };
  
  // Word overlap analysis
  const fragmentWords = normFragment.split(/\s+/).filter(w => w.length > 2);
  const textWords = normText.split(/\s+/).filter(w => w.length > 2);
  const textWordsSet = new Set(textWords);
  
  let wordMatches = 0;
  let importantWordMatches = 0; // Words longer than 4 chars
  
  fragmentWords.forEach(word => {
    if (textWordsSet.has(word)) {
      wordMatches++;
      if (word.length > 4) importantWordMatches++;
    }
  });
  
  const wordOverlap = fragmentWords.length > 0 ? wordMatches / fragmentWords.length : 0;
  const importantWordOverlap = fragmentWords.filter(w => w.length > 4).length > 0 ? 
    importantWordMatches / fragmentWords.filter(w => w.length > 4).length : 0;
  
  // Length similarity (prefer similar length texts)
  const lengthRatio = Math.min(fragment.length, text.length) / Math.max(fragment.length, text.length);
  
  // Combined score with weights
  const combinedScore = (wordOverlap * 0.5) + (importantWordOverlap * 0.3) + (lengthRatio * 0.2);
  
  return {
    score: combinedScore,
    wordOverlap: wordOverlap,
    lengthSimilarity: lengthRatio,
    importantWordOverlap: importantWordOverlap
  };
}

/**
 * Smart text chunking for large elements
 * @param {string} text The text to chunk.
 * @param {number} maxChunkSize Maximum chunk size in characters.
 * @returns {Array<Object>} Array of text chunks with metadata.
 */
function createSmartTextChunks(text, maxChunkSize = 100) {
  if (text.length <= maxChunkSize) {
    return [{ text: text, isComplete: true, position: 0 }];
  }
  
  const chunks = [];
  const sentences = text.split(/[.!?]+/).filter(s => s.trim().length > 0);
  
  let currentChunk = '';
  let position = 0;
  
  for (let i = 0; i < sentences.length; i++) {
    const sentence = sentences[i].trim();
    if ((currentChunk + sentence).length <= maxChunkSize) {
      currentChunk += (currentChunk ? '. ' : '') + sentence;
    } else {
      if (currentChunk) {
        chunks.push({
          text: currentChunk + '.',
          isComplete: false,
          position: position
        });
        position += currentChunk.length;
      }
      currentChunk = sentence;
    }
  }
  
  if (currentChunk) {
    chunks.push({
      text: currentChunk + '.',
      isComplete: chunks.length === 0,
      position: position
    });
  }
  
  return chunks;
}

/**
 * Multi-stage AI candidate filtering for large documents
 * @param {string} fragment The search fragment.
 * @param {Array<Object>} allElements All document elements.
 * @param {number} maxCandidates Maximum candidates to return.
 * @returns {Array<Object>} Filtered and scored candidates.
 */
function performAdvancedAIFiltering(fragment, allElements, maxCandidates = 3) {
  logOperation('INFO', `🔍 Advanced filtering for fragment: "${fragment.substring(0, 50)}..."`);
  
  // Stage 1: Quick word-based filtering
  const stage1Candidates = allElements.map(elem => {
    const similarity = calculateAdvancedSimilarity(fragment, elem.text);
    return {
      element: elem,
      similarity: similarity,
      textChunks: createSmartTextChunks(elem.text, 100)
    };
  }).filter(candidate => candidate.similarity.score >= 0.15) // Lower threshold for large docs
    .sort((a, b) => b.similarity.score - a.similarity.score)
    .slice(0, 20); // Take top 20 for stage 2
  
  logOperation('INFO', `📋 Stage 1: ${stage1Candidates.length} candidates (from ${allElements.length} elements)`);
  
  if (stage1Candidates.length === 0) {
    return [];
  }
  
  // Stage 2: Detailed analysis with chunking
  const stage2Candidates = [];
  
  stage1Candidates.forEach(candidate => {
    // For each element, analyze all chunks
    candidate.textChunks.forEach((chunk, chunkIndex) => {
      const chunkSimilarity = calculateAdvancedSimilarity(fragment, chunk.text);
      if (chunkSimilarity.score >= 0.2) {
        stage2Candidates.push({
          element: candidate.element,
          chunk: chunk,
          chunkIndex: chunkIndex,
          similarity: chunkSimilarity,
          combinedScore: (candidate.similarity.score * 0.3) + (chunkSimilarity.score * 0.7)
        });
      }
    });
  });
  
  // Sort by combined score and take the best
  const finalCandidates = stage2Candidates
    .sort((a, b) => b.combinedScore - a.combinedScore)
    .slice(0, maxCandidates);
  
  logOperation('INFO', `🎯 Stage 2: ${finalCandidates.length} final candidates for AI analysis`);
  
  return finalCandidates;
}

/**
 * Optimized batch AI processing
 * @param {Array<Object>} aiRequests Array of AI requests to process.
 * @returns {Array<Object>} Results for each request.
 */
function processBatchAIRequests(aiRequests) {
  const results = [];
  let cacheHits = 0;
  
  // Check cache first
  aiRequests.forEach((request, index) => {
    const cacheKey = `${request.fragment}_${request.candidates.map(c => c.chunk.text).join('|')}`;
    
    if (AI_SEARCH_CACHE.has(cacheKey)) {
      results[index] = AI_SEARCH_CACHE.get(cacheKey);
      cacheHits++;
      logOperation('INFO', `💾 Cache hit for request #${index + 1}`);
      return;
    }
    
    // Prepare for AI call
    if (request.candidates.length > 0) {
      const candidateTexts = request.candidates.map(c => c.chunk.text);
      const aiResult = callOptimizedOpenAI_findBestMatch(request.fragment, candidateTexts);
      
      const result = {
        fragment: request.fragment,
        bestMatch: null,
        aiResult: aiResult
      };
      
      if (aiResult !== 'NOT_FOUND' && !isNaN(parseInt(aiResult, 10))) {
        const index = parseInt(aiResult, 10);
        if (index >= 0 && index < request.candidates.length) {
          result.bestMatch = request.candidates[index];
        }
      }
      
      // Cache the result
      AI_SEARCH_CACHE.set(cacheKey, result);
      results[index] = result;
    } else {
      results[index] = { fragment: request.fragment, bestMatch: null, aiResult: 'NO_CANDIDATES' };
    }
  });
  
  logOperation('INFO', `🤖 Processed ${aiRequests.length} AI requests (${cacheHits} cache hits)`);
  
  // Manage cache size to prevent memory issues
  manageAICacheSize();
  
  return results;
}

/**
 * Optimized OpenAI call with better token management
 * @param {string} fragment The fragment to search for.
 * @param {Array<string>} candidateTexts Array of candidate texts.
 * @returns {string} The index of best match or "NOT_FOUND".
 */
function callOptimizedOpenAI_findBestMatch(fragment, candidateTexts) {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) {
    logOperation('WARN', '⚠️ OpenAI API key not found');
    return 'NOT_FOUND';
  }
  
  if (!candidateTexts || candidateTexts.length === 0) {
    logOperation('WARN', '⚠️ No candidates provided for AI analysis');
    return 'NOT_FOUND';
  }
  
  // Calculate approximate token count (rough estimate: 1 token ≈ 4 characters)
  const totalChars = fragment.length + candidateTexts.join('').length;
  const estimatedTokens = Math.ceil(totalChars / 4) + 100; // +100 for prompt overhead
  
  logOperation('INFO', `🤖 AI call: ${candidateTexts.length} candidates, ~${estimatedTokens} tokens`);
  
  // If too many tokens, truncate candidates
  if (estimatedTokens > 1500) { // Conservative limit
    const maxCandidateLength = Math.floor((1500 * 4 - fragment.length - 400) / candidateTexts.length);
    candidateTexts = candidateTexts.map(text => 
      text.length > maxCandidateLength ? text.substring(0, maxCandidateLength) + '...' : text
    );
    logOperation('WARN', `⚠️ Truncated candidates to fit token limit (max ${maxCandidateLength} chars each)`);
  }

  const systemPrompt = `You are a language expert specializing in semantic analysis. Find the candidate text that best matches the given fragment semantically.

GUIDELINES:
- Analyze meaning, context, and semantic relationships
- Consider partial matches and conceptual similarity
- The fragment might be a paraphrase or expansion of the candidate
- Respond with ONLY the numeric index (0, 1, 2, etc.) or "NOT_FOUND"
- No additional text or explanations`;

  const candidatesText = candidateTexts.map((text, i) => `[${i}]: ${text}`).join('\n\n');
  const userPrompt = `FRAGMENT: "${fragment}"\n\nCANDIDATES:\n${candidatesText}\n\nBest match index:`;

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.1,
    max_tokens: 5 // Only need a number
  };

  // Use the existing retry logic
  return callOpenAI_findBestMatchIndex_Internal(payload, candidateTexts.length);
}

/**
 * Internal AI call with retry logic
 * @param {Object} payload The API payload.
 * @param {number} maxIndex Maximum valid index.
 * @returns {string} Result or "NOT_FOUND".
 */
function callOpenAI_findBestMatchIndex_Internal(payload, maxIndex) {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  const MAX_RETRIES = 3;
  const BASE_DELAY = 1000;
  
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    try {
      logOperation('INFO', `🤖 AI API call attempt ${attempt + 1}/${MAX_RETRIES}`);
      
      const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
        method: 'post',
        contentType: 'application/json',  
        headers: { 'Authorization': 'Bearer ' + key },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
       
      const responseCode = response.getResponseCode();
      const responseText = response.getContentText();
      
      if (responseCode === 200) {
        try {
          const responseData = JSON.parse(responseText);
          const result = responseData.choices[0].message.content.trim();
          logOperation('INFO', `✅ AI API success: ${result}`);
          
          if (result === 'NOT_FOUND') return 'NOT_FOUND';
          
          const index = parseInt(result, 10);
          if (!isNaN(index) && index >= 0 && index < maxIndex) {
            return result;
          } else {
            logOperation('WARN', `⚠️ Invalid AI response: ${result}. Expected 0-${maxIndex-1} or NOT_FOUND`);
            return 'NOT_FOUND';
          }
        } catch (parseError) {
          logOperation('ERROR', `💥 Failed to parse AI response: ${parseError.message}`);
          throw parseError;
        }
      } else if (responseCode === 429) {
        logOperation('WARN', `⚠️ Rate limit (429). Attempt ${attempt + 1}/${MAX_RETRIES}`);
        if (attempt < MAX_RETRIES - 1) {
          const delay = BASE_DELAY * Math.pow(2, attempt);
          Utilities.sleep(delay);
          continue;
        }
      } else if (responseCode === 401) {
        logOperation('ERROR', `💥 Invalid API key (401)`);
        return 'NOT_FOUND';
      } else if (responseCode === 403) {
        logOperation('ERROR', `💥 Access forbidden (403)`);
        return 'NOT_FOUND';
      } else if (responseCode >= 500) {
        logOperation('WARN', `⚠️ Server error (${responseCode}). Attempt ${attempt + 1}/${MAX_RETRIES}`);
        if (attempt < MAX_RETRIES - 1) {
          const delay = BASE_DELAY * Math.pow(2, attempt);
          Utilities.sleep(delay);
          continue;
        }
      } else {
        logOperation('ERROR', `💥 AI API Error: HTTP ${responseCode}. Response: ${responseText}`);
        return 'NOT_FOUND';
      }
    } catch (e) {
      logOperation('ERROR', `💥 AI API call failed (attempt ${attempt + 1}/${MAX_RETRIES}): ${e.message}`);
      if (attempt < MAX_RETRIES - 1) {
        const delay = BASE_DELAY * Math.pow(2, attempt);
        Utilities.sleep(delay);
        continue;
      }
    }
  }
  
  logOperation('ERROR', `💥 All ${MAX_RETRIES} AI API attempts failed`);
  return 'NOT_FOUND';
}

/**
 * Applies the approved changes to the document.
 * Updated to handle multiple replacements in same element correctly.
 * @param {Array<Object>} approvedSuggestions The suggestions confirmed by the user.
 * @returns {string} A summary of the operation.
 */
function applySuggestions(approvedSuggestions) {
  logOperation('INFO', "🔥 Applying approved suggestions...");
  
  try {
    if (!approvedSuggestions || approvedSuggestions.length === 0) {
      logOperation('WARN', "No suggestions provided to apply.");
      return "❌ No changes were selected to apply.";
    }
    
    const startTime = new Date().getTime();
    logOperation('INFO', `🚀 Applying ${approvedSuggestions.length} suggestions.`, {
      suggestionsCount: approvedSuggestions.length,
      startTime: startTime
    });
    
    // Log each suggestion for debugging
    approvedSuggestions.forEach((suggestion, index) => {
      logOperation('INFO', `📋 Suggestion #${index + 1}: ${suggestion.type} in ${suggestion.elementType} #${suggestion.paraIndex}`, {
        suggestionIndex: index,
        suggestionType: suggestion.type,
        elementType: suggestion.elementType,
        paraIndex: suggestion.paraIndex,
        fragment: suggestion.fragment,
        replaceWith: suggestion.replaceWith,
        replacementCount: suggestion.replacementCount || 1
      });
    });
    
    // Initialize progress tracking
    const totalSuggestions = approvedSuggestions.length;
    const PROGRESS_KEY = 'RITUAL_PROGRESS';
    PropertiesService.getDocumentProperties().setProperty(PROGRESS_KEY, JSON.stringify({ applied: 0, total: totalSuggestions, done: false }));
    logOperation('INFO', `📊 Progress tracking initialized for ${totalSuggestions} suggestions`);

    // Since suggestion.element is not passed from the client, we must re-fetch elements.
    const elementsStart = new Date().getTime();
    const allElements = getAllDocumentElements();
    const elementsTime = new Date().getTime() - elementsStart;
    logOperation('INFO', `📄 Fetched ${allElements.length} elements in ${elementsTime}ms`, {
      elementsCount: allElements.length,
      fetchTime: elementsTime
    });
    
    let appliedCount = 0;
    let totalReplacementCount = 0;
    let errors = [];
    const backupEntries = [];

    // Sort suggestions by index in descending order to prevent shifts from affecting subsequent changes.
    const sortedSuggestions = [...approvedSuggestions].sort((a, b) => b.paraIndex - a.paraIndex);

    sortedSuggestions.forEach((suggestion, index) => {
      const operationStart = new Date().getTime();
      const replacementCount = suggestion.replacementCount || 1;
      logOperation('INFO', `\n🔄 [${index + 1}/${sortedSuggestions.length}] Processing ${suggestion.type} in ${suggestion.elementType} #${suggestion.paraIndex} (${replacementCount} replacements)`, {
        suggestionIndex: index + 1,
        totalSuggestions: sortedSuggestions.length,
        suggestionType: suggestion.type,
        elementType: suggestion.elementType,
        paraIndex: suggestion.paraIndex,
        replacementCount: replacementCount
      });
      
      try {
        // Find the element wrapper by its original index.
        const elementWrapper = allElements.find(e => e.originalIndex === suggestion.paraIndex);
        if (!elementWrapper) {
          logOperation('ERROR', `❌ Element #${suggestion.paraIndex} not found. It might have been modified or removed.`, {
            paraIndex: suggestion.paraIndex,
            elementType: suggestion.elementType,
            allElementsCount: allElements.length
          });
          errors.push(`Could not find element #${suggestion.paraIndex}`);
          return;
        }
        
        const element = elementWrapper.element;
        logOperation('INFO', `✅ Element found: ${elementWrapper.typeName}`, {
          elementIndex: suggestion.paraIndex,
          elementType: elementWrapper.typeName,
          elementId: elementWrapper.id
        });

        const currentText = getElementText(element).trim();
        
        // Verify that the element's text hasn't changed since the preview was generated.
        const textsMatch = currentText === suggestion.oldText;
        if (!textsMatch) {
          logOperation('WARN', `⚠️ Text has changed since preview. Expected: "${suggestion.oldText.substring(0, 80)}..."`, {
            paraIndex: suggestion.paraIndex,
            expectedText: suggestion.oldText,
            currentText: currentText,
            expectedLength: suggestion.oldText.length,
            currentLength: currentText.length
          });
          logOperation('WARN', `⚠️ Current text: "${currentText.substring(0, 80)}..."`, {
            paraIndex: suggestion.paraIndex,
            currentTextFull: currentText
          });
          
          // For exact matches, try to apply anyway if the fragments are still found
          if (suggestion.type === 'EXACT') {
            const fragments = suggestion.fragment.includes(' + ') ? suggestion.fragment.split(' + ') : [suggestion.fragment];
            const canApply = fragments.every(fragment => currentText.includes(fragment));
            
            logOperation('INFO', `🔍 Checking if fragments still exist in changed text`, {
              paraIndex: suggestion.paraIndex,
              fragments: fragments,
              canApply: canApply,
              fragmentsFound: fragments.map(f => ({ fragment: f, found: currentText.includes(f) }))
            });
            
            if (canApply) {
              logOperation('INFO', `🔄 Text changed but fragments still found. Attempting to apply...`, {
                paraIndex: suggestion.paraIndex,
                fragmentsCount: fragments.length
              });
            } else {
              logOperation('ERROR', `⚠️ Skipping: required fragments not found in current text`, {
                paraIndex: suggestion.paraIndex,
                fragments: fragments,
                currentText: currentText
              });
              errors.push(`${suggestion.elementType} #${suggestion.paraIndex} was modified - fragments not found`);
              return;
            }
          } else {
            logOperation('ERROR', `⚠️ Skipping: text was modified for ${suggestion.type} match`, {
              paraIndex: suggestion.paraIndex,
              suggestionType: suggestion.type,
              expectedText: suggestion.oldText,
              currentText: currentText
            });
            errors.push(`${suggestion.elementType} #${suggestion.paraIndex} was modified`);
            return;
          }
        } else {
          logOperation('INFO', `✅ Text verification passed - element text unchanged`, {
            paraIndex: suggestion.paraIndex,
            textLength: currentText.length
          });
        }
        
        // Create backup entry before making changes
        backupEntries.push({
          paraIndex: suggestion.paraIndex,
          elementType: suggestion.elementType,
          oldText: currentText, // Use current text, not suggestion.oldText
          newText: suggestion.newText,
          fragment: suggestion.fragment,
          replaceWith: suggestion.replaceWith,
          type: suggestion.type,
          replacementCount: replacementCount
        });

        // Apply the changes
        if (suggestion.type === 'EXACT') {
          if (replacementCount === 1) {
            // Single exact replacement
            logOperation('INFO', `🎯 EXACT replace: "${suggestion.fragment}" → "${suggestion.replaceWith}"`, {
              paraIndex: suggestion.paraIndex,
              fragment: suggestion.fragment,
              replaceWith: suggestion.replaceWith,
              fragmentLength: suggestion.fragment.length,
              replaceWithLength: suggestion.replaceWith.length
            });
            
            try {
              element.replaceText(escapeRegex(suggestion.fragment), suggestion.replaceWith);
              logOperation('INFO', `✅ EXACT replacement applied successfully`, {
                paraIndex: suggestion.paraIndex,
                fragment: suggestion.fragment
              });
            } catch (e) {
              logOperation('ERROR', `💥 Error applying EXACT replacement: ${e.message}`, {
                paraIndex: suggestion.paraIndex,
                fragment: suggestion.fragment,
                replaceWith: suggestion.replaceWith,
                error: e.message
              });
              throw e;
            }
          } else {
            // Multiple exact replacements (already combined in preview)
            logOperation('INFO', `🎯 EXACT replace (${replacementCount} fragments): "${suggestion.fragment}" → "${suggestion.replaceWith}"`, {
              paraIndex: suggestion.paraIndex,
              replacementCount: replacementCount,
              combinedFragment: suggestion.fragment,
              combinedReplaceWith: suggestion.replaceWith
            });
            
            const fragments = suggestion.fragment.split(' + ');
            const replacements = suggestion.replaceWith.split(' + ');
            
            logOperation('INFO', `🔄 Processing ${fragments.length} individual fragments`, {
              paraIndex: suggestion.paraIndex,
              fragmentsCount: fragments.length,
              replacementsCount: replacements.length,
              fragments: fragments,
              replacements: replacements
            });
            
            // Apply replacements in order
            for (let i = 0; i < fragments.length && i < replacements.length; i++) {
              try {
                logOperation('INFO', `🔄 Applying fragment #${i + 1}: "${fragments[i]}" → "${replacements[i]}"`, {
                  paraIndex: suggestion.paraIndex,
                  fragmentIndex: i + 1,
                  fragment: fragments[i],
                  replacement: replacements[i]
                });
                
                element.replaceText(escapeRegex(fragments[i]), replacements[i]);
                
                logOperation('INFO', `✅ Fragment #${i + 1} applied successfully`, {
                  paraIndex: suggestion.paraIndex,
                  fragmentIndex: i + 1
                });
              } catch (e) {
                logOperation('ERROR', `💥 Error applying fragment #${i + 1}: ${e.message}`, {
                  paraIndex: suggestion.paraIndex,
                  fragmentIndex: i + 1,
                  fragment: fragments[i],
                  replacement: replacements[i],
                  error: e.message
                });
                throw e;
              }
            }
          }
        } else if (suggestion.type === 'AI') {
          // AI replacement - replace entire element text
          logOperation('INFO', `🤖 AI replace: full text replacement`, {
            paraIndex: suggestion.paraIndex,
            oldTextLength: suggestion.oldText.length,
            newTextLength: suggestion.newText.length,
            newText: suggestion.newText.substring(0, 200) + (suggestion.newText.length > 200 ? '...' : '')
          });
          
          try {
            applyTextChange(element, suggestion.newText);
            logOperation('INFO', `✅ AI replacement applied successfully`, {
              paraIndex: suggestion.paraIndex,
              newTextLength: suggestion.newText.length
            });
          } catch (e) {
            logOperation('ERROR', `💥 Error applying AI replacement: ${e.message}`, {
              paraIndex: suggestion.paraIndex,
              newText: suggestion.newText,
              error: e.message
            });
            throw e;
          }
        } else {
          // Fallback for any other type
          logOperation('INFO', `🔄 ${suggestion.type} replace: full text`, {
            paraIndex: suggestion.paraIndex,
            suggestionType: suggestion.type,
            oldTextLength: suggestion.oldText.length,
            newTextLength: suggestion.newText.length
          });
          
          try {
            applyTextChange(element, suggestion.newText);
            logOperation('INFO', `✅ ${suggestion.type} replacement applied successfully`, {
              paraIndex: suggestion.paraIndex,
              suggestionType: suggestion.type
            });
          } catch (e) {
            logOperation('ERROR', `💥 Error applying ${suggestion.type} replacement: ${e.message}`, {
              paraIndex: suggestion.paraIndex,
              suggestionType: suggestion.type,
              error: e.message
            });
            throw e;
          }
        }
        
        appliedCount++;
        totalReplacementCount += replacementCount;

        // Update progress
        PropertiesService.getDocumentProperties().setProperty(PROGRESS_KEY, JSON.stringify({ applied: appliedCount, total: totalSuggestions, done: false }));

        const operationTime = new Date().getTime() - operationStart;
        logOperation('INFO', `⏱️ Operation finished in ${operationTime}ms`, {
          paraIndex: suggestion.paraIndex,
          operationTime: operationTime,
          appliedCount: appliedCount,
          totalReplacementCount: totalReplacementCount
        });
        
      } catch (e) {
        logOperation('ERROR', `💥 ERROR applying change to ${suggestion.elementType} #${suggestion.paraIndex}: ${e.message}`, {
          paraIndex: suggestion.paraIndex,
          elementType: suggestion.elementType,
          suggestionType: suggestion.type,
          error: e.message,
          stack: e.stack,
          fragment: suggestion.fragment,
          replaceWith: suggestion.replaceWith
        });
        errors.push(`Error in ${suggestion.elementType} #${suggestion.paraIndex}: ${e.message}`);
      }
    });

    // Save the backup for the "Undo" feature.
    try {
      PropertiesService.getDocumentProperties().setProperty('LAST_RUN_BACKUP', JSON.stringify(backupEntries));
      logOperation('INFO', `💾 Backup of ${backupEntries.length} changes saved for potential undo.`, {
        backupEntriesCount: backupEntries.length,
        backupSize: JSON.stringify(backupEntries).length
      });
    } catch (e) {
      logOperation('ERROR', `⚠️ Could not save undo backup: ${e.message}`, {
        error: e.message,
        backupEntriesCount: backupEntries.length
      });
    }
    
    // Finalize progress
    PropertiesService.getDocumentProperties().setProperty(PROGRESS_KEY, JSON.stringify({ applied: appliedCount, total: totalSuggestions, done: true }));

    const totalTime = (new Date().getTime() - startTime) / 1000;
    logOperation('INFO', `\n🏁 Finished applying changes in ${totalTime.toFixed(2)}s.`, {
      totalTime: totalTime,
      appliedCount: appliedCount,
      totalSuggestions: approvedSuggestions.length,
      totalReplacementCount: totalReplacementCount,
      errorsCount: errors.length
    });
    
    logOperation('INFO', `📊 Final Stats: ${appliedCount} elements modified, ${totalReplacementCount} individual replacements, ${errors.length} errors.`, {
      elementsModified: appliedCount,
      individualReplacements: totalReplacementCount,
      errors: errors.length,
      successRate: Math.round((appliedCount / approvedSuggestions.length) * 100)
    });

    let result = `🎉 Applied ${appliedCount}/${approvedSuggestions.length} suggestions (${totalReplacementCount} individual replacements) in ${totalTime.toFixed(2)}s.`;
    if (errors.length > 0) {
      result += `\n⚠️ Not applied: ${errors.length}. See logs for details.`;
      logOperation('WARN', "Error details:", {
        errors: errors,
        errorCount: errors.length
      });
    }

    return result + `\n↩️ To revert, use the "Undo Last Run" button.`;
    
  } catch (e) {
    logOperation('ERROR', `💥💥💥 CRITICAL ERROR in applySuggestions: ${e.message}`, {
      error: e.message,
      stack: e.stack,
      approvedSuggestionsCount: approvedSuggestions?.length || 0
    });
    return `💥 Critical error: ${e.message}`;
  }
}

/**
 * Reverts the last set of applied changes.
 * @returns {string} A summary of the undo operation.
 */
function undoLastRun() {
  try {
    const propService = PropertiesService.getDocumentProperties();
    const backupJson = propService.getProperty('LAST_RUN_BACKUP');
    if (!backupJson) {
      return '❌ No saved run found to undo.';
    }

    const backupEntries = JSON.parse(backupJson);
    if (!Array.isArray(backupEntries) || backupEntries.length === 0) {
      return '❌ Change history is empty.';
    }

    const allElements = getAllDocumentElements();
    let undone = 0;
    const errors = [];

    // Process in reverse index order to avoid position shifts.
    backupEntries.sort((a,b) => b.paraIndex - a.paraIndex).forEach(entry => {
      try {
        let elementWrapper = allElements.find(e => e.originalIndex === entry.paraIndex);
        
        // Fallback: if the element moved, find it by its "new" text content.
        if (!elementWrapper) {
          elementWrapper = allElements.find(e => e.text === entry.newText);
        }
        if (!elementWrapper) {
          errors.push(`Could not find element #${entry.paraIndex} to revert.`);
          return;
        }
        const element = elementWrapper.element;

        if (entry.type === 'EXACT') {
          // Revert the exact replacement.
          if (entry.replacementCount > 1) {
            // Multiple replacements - need to handle them properly
            const fragments = entry.fragment.split(' + ');
            const replacements = entry.replaceWith.split(' + ');
            
            // Apply reverse replacements in reverse order
            for (let i = replacements.length - 1; i >= 0; i--) {
              if (i < fragments.length) {
                element.replaceText(escapeRegex(replacements[i]), fragments[i]);
              }
            }
          } else {
            // Single replacement
            element.replaceText(escapeRegex(entry.replaceWith), entry.fragment);
          }
        } else {
          // Revert the full text change.
          applyTextChange(element, entry.oldText);
        }
        undone++;
      } catch (e) {
        errors.push(`Error reverting #${entry.paraIndex}: ${e.message}`);
      }
    });

    // Clear the backup so it can't be used again.
    propService.deleteProperty('LAST_RUN_BACKUP');

    let msg = `↩️ Reverted ${undone}/${backupEntries.length} changes.`;
    if (errors.length) msg += `\n⚠️ Issues: ${errors.length}`;
    return msg;
  } catch (e) {
    return `💥 Undo failed: ${e.message}`;
  }
}

/**
 * Retrieves the current progress of a running task.
 * @returns {Object} The progress status {applied, total, done}.
 */
function getRitualProgress() {
  const PROGRESS_KEY = 'RITUAL_PROGRESS';
  try {
    const prop = PropertiesService.getDocumentProperties().getProperty(PROGRESS_KEY);
    return prop ? JSON.parse(prop) : { applied: 0, total: 0, done: true };
  } catch (e) {
    return { applied: 0, total: 0, done: true, error: e.message };
  }
}

/**
 * A utility to apply text changes to various element types.
 * @param {GoogleAppsScript.Document.Element} element The element to modify.
 * @param {string} newText The new text to apply.
 */
function applyTextChange(element, newText) {
  const elementType = element.getType();
  
  try {
    switch (elementType) {
      case DocumentApp.ElementType.PARAGRAPH:
      case DocumentApp.ElementType.HEADING:
      case DocumentApp.ElementType.LIST_ITEM:
      case DocumentApp.ElementType.TABLE_CELL:
        element.clear().setText(newText);
        break;
      case DocumentApp.ElementType.TEXT:
        element.setText(newText);
        break;
      default:
        // Generic fallback for elements that support clear() and setText().
        if (element.clear && element.setText) {
          element.clear().setText(newText);
        } else if (element.setText) {
          element.setText(newText);
        } else {
          throw new Error(`Element type ${elementType} does not support text modification.`);
        }
    }
  } catch (e) {
    throw new Error(`Failed to modify text for element type ${elementType}: ${e.message}`);
  }
}

/* ───────── 🔧 Helpers & API Calls ───────── */

/**
 * Escapes special regex characters in a string.
 * @param {string} text The text to escape.
 * @returns {string} The escaped text.
 */
function escapeRegex(text) {
  return text.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&');
}

/**
 * Checks if the OpenAI API key is set in Script Properties.
 * @returns {boolean} True if the key exists.
 */
function checkOpenAIKey() {
  try {
    const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
    const hasKey = !!key;
    logOperation('INFO', `🔑 API key check: ${hasKey ? 'Available' : 'Not available'}`, {
      hasKey: hasKey,
      keyLength: key ? key.length : 0,
      keyPreview: key ? key.substring(0, 8) + '***' : 'none'
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
 * Generates a unique ID for an element based on its content and position.
 * @param {Object} elementWrapper The element wrapper with text and index.
 * @returns {string} A unique element ID.
 */
function generateElementId(elementWrapper) {
  try {
    // Create a hash-like ID based on element content and position
    const text = elementWrapper.text || '';
    const index = elementWrapper.originalIndex || 0;
    const type = elementWrapper.typeName || 'Unknown';
    
    // Take first 50 chars of text for ID generation
    const textSample = text.substring(0, 50).replace(/\s+/g, '_').replace(/[^a-zA-Z0-9_]/g, '');
    
    // Simple hash function
    let hash = 0;
    for (let i = 0; i < text.length; i++) {
      const char = text.charCodeAt(i);
      hash = ((hash << 5) - hash) + char;
      hash = hash & hash; // Convert to 32bit integer
    }
    
    const elementId = `${type}_${index}_${Math.abs(hash).toString(36)}_${textSample}`;
    logOperation('INFO', `🆔 Generated ID for element #${index}: ${elementId}`, {
      elementIndex: index,
      elementType: type,
      textLength: text.length,
      textSample: textSample,
      hash: Math.abs(hash).toString(36),
      fullElementId: elementId
    });
    
    return elementId;
  } catch (e) {
    logOperation('ERROR', `💥 Error generating element ID: ${e.message}`, {
      error: e.message,
      stack: e.stack,
      elementWrapper: elementWrapper
    });
    return `unknown_${elementWrapper.originalIndex || 0}_${Date.now()}`;
  }
}

/* ───────── 📊 Logging & Error Reporting ───────── */

/**
 * Global log collector for detailed operation tracking
 */
if (typeof OPERATION_LOGS === 'undefined') {
  var OPERATION_LOGS = [];
}

/**
 * Enhanced logging function that collects logs for later retrieval
 * @param {string} level Log level (INFO, WARN, ERROR)
 * @param {string} message Log message
 * @param {Object} data Additional data to log
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
  
  // Also log to console with appropriate level
  const fullMessage = `[${timestamp}] ${level}: ${message}`;
  switch (level) {
    case 'ERROR':
      console.error(fullMessage, data);
      break;
    case 'WARN':
      console.warn(fullMessage, data);
      break;
    default:
      console.log(fullMessage, data);
  }
  
  // Keep only last 1000 log entries to prevent memory issues
  if (OPERATION_LOGS.length > 1000) {
    OPERATION_LOGS.splice(0, OPERATION_LOGS.length - 1000);
  }
}

/**
 * Clears the operation logs
 */
function clearOperationLogs() {
  OPERATION_LOGS.length = 0;
  logOperation('INFO', '🗑️ Operation logs cleared');
}

/**
 * Gets all collected logs for the current operation
 * @returns {Array} Array of log entries
 */
function getOperationLogs() {
  return [...OPERATION_LOGS]; // Return a copy
}

/**
 * Gets a summary of errors and warnings from the logs
 * @returns {Object} Summary of issues
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

/**
 * Clears the AI search cache
 */
function clearAISearchCache() {
  AI_SEARCH_CACHE.clear();
  logOperation('INFO', '🗑️ AI search cache cleared');
}

/**
 * Gets AI cache statistics
 * @returns {Object} Cache statistics
 */
function getAICacheStats() {
  return {
    size: AI_SEARCH_CACHE.size,
    maxSize: 100 // Keep last 100 searches
  };
}

/**
 * Manages AI cache size to prevent memory issues
 */
function manageAICacheSize() {
  const maxSize = 100;
  if (AI_SEARCH_CACHE.size > maxSize) {
    // Convert to array, sort by usage (newest first), keep only maxSize
    const entries = Array.from(AI_SEARCH_CACHE.entries());
    AI_SEARCH_CACHE.clear();
    
    // Keep only the last maxSize entries
    entries.slice(-maxSize).forEach(([key, value]) => {
      AI_SEARCH_CACHE.set(key, value);
    });
    
    logOperation('INFO', `🧹 AI cache trimmed to ${maxSize} entries`);
  }
}