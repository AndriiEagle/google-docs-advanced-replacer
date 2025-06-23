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
* 4. Multi-level matching: EXACT, FUZZY (similarity-based), and AI-powered.
* 5. "Undo" functionality to revert the last batch of changes.
* 6. Real-time progress bar for large documents.
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
  const doc = DocumentApp.getActiveDocument();
  if (!doc) throw new Error('No active Google Docs document found.');
  const body = doc.getBody();
  if (!body) throw new Error('Could not retrieve the document body.');
  return body;
}

/**
 * Retrieves all processable text-containing elements from the document.
 * This function recursively traverses the document structure.
 * @returns {Array<Object>} An array of element wrappers.
 */
function getAllDocumentElements() {
  const body = safeGetBody();
  const elements = [];
  let index = 0;

  function processElement(element) {
    const elementType = element.getType();
    const text = getElementText(element).trim();
    
    // FIX v5.1.1: Skip the root BODY_SECTION to avoid matching the entire document as one element.
    if (elementType !== DocumentApp.ElementType.BODY_SECTION && text.length > 0) {
      elements.push({
        element: element,
        text: text,
        type: elementType,
        typeName: getElementTypeName(elementType),
        originalIndex: index++
      });
    }

    // Recursively process children of container elements.
    if (element.getNumChildren && element.getNumChildren() > 0) {
      for (let i = 0; i < element.getNumChildren(); i++) {
        try {
          const child = element.getChild(i);
          processElement(child);
        } catch (e) {
          // Log and skip children that can't be processed (e.g., permissions issues).
          console.warn(`Skipping a child element due to error: ${e.message}`);
        }
      }
    }
  }

  processElement(body);
  console.log(`Found ${elements.length} elements of various types.`);
  return elements;
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
 * Advanced text normalization for better fuzzy matching.
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

/**
 * Calculates a similarity score between two strings, weighted by element type.
 * @param {string} fragment The search fragment.
 * @param {string} text The text to compare against.
 * @param {GoogleAppsScript.Document.ElementType} elementType The type of the element containing the text.
 * @returns {number} A similarity score between 0 and 1.
 */
function calculateSimilarity(fragment, text, elementType) {
  const normFragment = normalizeText(fragment);
  const normText = normalizeText(text);
  
  if (!normFragment || !normText) return 0;
  if (normFragment === normText) return 1;
  
  const maxLen = Math.max(normFragment.length, normText.length);  
  if (maxLen === 0) return 1;
  const baseSimilarity = 1 - levenshtein(normFragment, normText) / maxLen;
  
  // Weighting boosts for more important element types.
  let typeBonus = 1.0;
  switch (elementType) {
    case DocumentApp.ElementType.HEADING:
      typeBonus = 1.3; // Headings are most important.
      break;
    case DocumentApp.ElementType.LIST_ITEM:
      typeBonus = 1.1; // List items are important.
      break;
    case DocumentApp.ElementType.PARAGRAPH:
      typeBonus = 1.0; // Standard text.
      break;
    default:
      typeBonus = 0.9; // Lower priority for other types.
  }
  
  // Bonus for exact word matches within the text.
  const fragmentWords = normFragment.split(/\s+/).filter(w => w.length > 2);
  const textWords = new Set(normText.split(/\s+/));
  let exactMatches = 0;
  
  fragmentWords.forEach(fragWord => {
    if (textWords.has(fragWord)) {
      exactMatches++;
    }
  });
  
  const wordBonus = fragmentWords.length > 0 ? (exactMatches / fragmentWords.length) * 0.2 : 0;
  
  return Math.min(1.0, baseSimilarity * typeBonus + wordBonus);
}

/**
 * Standard Levenshtein distance implementation.
 * @param {string} a First string.
 * @param {string} b Second string.
 * @returns {number} The Levenshtein distance.
 */
function levenshtein(a = '', b = '') {
  if (a.length === 0) return b.length;
  if (b.length === 0) return a.length;
  const matrix = [];
  for (let i = 0; i <= b.length; i++) { matrix[i] = [i]; }
  for (let j = 0; j <= a.length; j++) { matrix[0][j] = j; }
  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      const cost = b.charAt(i - 1) == a.charAt(j - 1) ? 0 : 1;
      matrix[i][j] = Math.min(
        matrix[i - 1][j - 1] + cost,
        matrix[i][j - 1] + 1,
        matrix[i - 1][j] + 1
      );
    }
  }
  return matrix[b.length][a.length];
}

/* ───────── 🚀 Main Processing Engine ───────── */

/**
 * Generates replacement previews based on a single pass over the document.
 * @param {Array<Object>} directives An array of {fragment, replaceWith} objects.
 * @param {Object} settings An object with {fuzzyThreshold, aiThreshold}.
 * @returns {Array<Object>} A sorted list of suggested changes.
 */
function generatePreview(directives, settings) {
  console.log(`🚀 Starting preview generation for ${directives.length} directives.`);
  
  const hasApiKey = checkOpenAIKey();
  const suggestions = [];
  
  // Get all document elements in one go.
  const allElements = getAllDocumentElements();
  
  directives.forEach((directive, directiveIndex) => {
    const { fragment, replaceWith } = directive;
    
    if (!fragment || !replaceWith) {
      console.warn(`⚠️ Directive #${directiveIndex + 1} skipped: empty fields.`);
      return;
    }

    console.log(`\n🎯 [Directive #${directiveIndex + 1}] Processing: "${fragment}" → "${replaceWith}"`);
    
    let bestExactMatch = null;
    let bestFuzzyMatch = null;
    let aiCandidates = [];
    
    // Single pass over all elements with prioritized logic.
    allElements.forEach(elem => {
      // PRIORITY 1: Look for exact matches.
      if (elem.text.includes(fragment)) {
        // A simple .includes() check is faster before trying heavier operations.
        try {
          // findText is only available on some elements, and it confirms the match.
          if (elem.element.findText && elem.element.findText(fragment)) {
            if (!bestExactMatch) { // Take the first exact match found.
              bestExactMatch = {
                element: elem,
                similarity: 1.0,
                newText: elem.text.replace(new RegExp(escapeRegex(fragment), 'g'), replaceWith),
                matchType: 'EXACT'
              };
            }
          }
        } catch (e) {
          // Element doesn't support findText, ignore.
        }
      }
      
      // PRIORITY 2: If no exact match, analyze for fuzzy matches.
      if (!bestExactMatch) {
        const similarity = calculateSimilarity(fragment, elem.text, elem.element.getType());
        
        if (similarity >= settings.fuzzyThreshold) {
          if (!bestFuzzyMatch || similarity > bestFuzzyMatch.similarity) {
            bestFuzzyMatch = {
              element: elem,
              similarity: similarity,
              newText: replaceWith, // For FUZZY, we replace the entire element's text.
              matchType: 'FUZZY'
            };
          }
        }
        
        // PRIORITY 3: If similarity is within AI threshold, collect as a candidate.
        if (similarity >= settings.aiThreshold) {
          aiCandidates.push({
            element: elem,
            similarity: similarity
          });
        }
      }
    });

    // Determine the best result based on the priority hierarchy.
    let bestMatch = bestExactMatch || bestFuzzyMatch;
    
    // AI analysis is the last resort if no other matches are found.
    if (!bestMatch && hasApiKey && aiCandidates.length > 0) {
      console.log(`   🤖 AI: Analyzing ${aiCandidates.length} candidates...`);
      
      aiCandidates.sort((a, b) => b.similarity - a.similarity);
      const topCandidates = aiCandidates.slice(0, 5); // Limit to top 5 for performance.
      
      const candidateTexts = topCandidates.map(c => c.element.text);
      const bestIndex = callOpenAI_findBestMatchIndex(fragment, candidateTexts);

      if (bestIndex !== 'NOT_FOUND' && !isNaN(parseInt(bestIndex, 10))) {
        const index = parseInt(bestIndex, 10);
        if (index >= 0 && index < topCandidates.length) {
          bestMatch = {
            element: topCandidates[index].element,
            similarity: 0.95, // Assign a high similarity score for AI matches.
            newText: replaceWith,
            matchType: 'AI'
          };
        }
      }
    }

    // If a match was found, add it to the suggestions list.
    if (bestMatch) {
      console.log(`   ✅ ${bestMatch.matchType}: Found in ${bestMatch.element.typeName} #${bestMatch.element.originalIndex} (Score: ${Math.round(bestMatch.similarity * 100)}%)`);
      
      suggestions.push({
        type: bestMatch.matchType,
        paraIndex: bestMatch.element.originalIndex,
        elementType: bestMatch.element.typeName,
        similarity: bestMatch.similarity,
        oldText: bestMatch.element.text,
        newText: bestMatch.newText,
        fragment: fragment,
        replaceWith: replaceWith,
        directiveIndex: directiveIndex,
        element: bestMatch.element.element // Pass the raw element for later use
      });
    } else {
      console.log(`   ❌ No suitable match found.`);
    }
  });

  suggestions.sort((a, b) => a.paraIndex - b.paraIndex);
  console.log(`\n🎉 Preview complete. Found ${suggestions.length}/${directives.length} suggestions.`);
  
  return suggestions;
}


/**
 * Applies the approved changes to the document.
 * @param {Array<Object>} approvedSuggestions The suggestions confirmed by the user.
 * @returns {string} A summary of the operation.
 */
function applySuggestions(approvedSuggestions) {
  console.log("🔥 Applying approved suggestions...");
  
  try {
    if (!approvedSuggestions || approvedSuggestions.length === 0) {
      console.log("No suggestions provided to apply.");
      return "❌ No changes were selected to apply.";
    }
    
    const startTime = new Date().getTime();
    console.log(`🚀 Applying ${approvedSuggestions.length} changes.`);
    
    // Initialize progress tracking
    const totalSuggestions = approvedSuggestions.length;
    const PROGRESS_KEY = 'RITUAL_PROGRESS';
    PropertiesService.getDocumentProperties().setProperty(PROGRESS_KEY, JSON.stringify({ applied: 0, total: totalSuggestions, done: false }));


    // Since suggestion.element is not passed from the client, we must re-fetch elements.
    const elementsStart = new Date().getTime();
    const allElements = getAllDocumentElements();
    console.log(`📄 Fetched ${allElements.length} elements in ${new Date().getTime() - elementsStart}ms`);
    
    let appliedCount = 0;
    let errors = [];
    const backupEntries = [];

    // Sort suggestions by index in descending order to prevent shifts from affecting subsequent changes.
    const sortedSuggestions = [...approvedSuggestions].sort((a, b) => b.paraIndex - a.paraIndex);

    sortedSuggestions.forEach((suggestion, index) => {
      const operationStart = new Date().getTime();
      console.log(`\n🔄 [${index + 1}/${sortedSuggestions.length}] Processing ${suggestion.type} in ${suggestion.elementType} #${suggestion.paraIndex}`);
      
      try {
        // Find the element wrapper by its original index.
        const elementWrapper = allElements.find(e => e.originalIndex === suggestion.paraIndex);
        if (!elementWrapper) {
          console.warn(`❌ Element #${suggestion.paraIndex} not found. It might have been modified or removed.`);
          errors.push(`Could not find element #${suggestion.paraIndex}`);
          return;
        }
        
        const element = elementWrapper.element;
        console.log(`✅ Element found: ${elementWrapper.typeName}`);

        const currentText = getElementText(element).trim();
        
        // Verify that the element's text hasn't changed since the preview was generated.
        const textsMatch = currentText === suggestion.oldText;
        if (!textsMatch) {
          console.log(`⚠️ Text has changed. Expected: "${suggestion.oldText.substring(0, 80)}..."`);
          console.log(`⚠️ Received: "${currentText.substring(0, 80)}..."`);
        }
        
        if (textsMatch) {
          // Add entry to backup before making the change.
          backupEntries.push({
            paraIndex: suggestion.paraIndex,
            elementType: suggestion.elementType,
            oldText: suggestion.oldText,
            newText: suggestion.newText,
            fragment: suggestion.fragment,
            replaceWith: suggestion.replaceWith,
            type: suggestion.type
          });

          if (suggestion.type === 'EXACT') {
            console.log(`🎯 EXACT replace: "${suggestion.fragment}" → "${suggestion.replaceWith}"`);
            element.replaceText(escapeRegex(suggestion.fragment), suggestion.replaceWith);
          } else {
            console.log(`🔄 ${suggestion.type} replace: full text`);
            applyTextChange(element, suggestion.newText);
          }
          
          appliedCount++;

          // Update progress
          PropertiesService.getDocumentProperties().setProperty(PROGRESS_KEY, JSON.stringify({ applied: appliedCount, total: totalSuggestions, done: false }));

        } else {
          console.warn(`⚠️ Skipped change because text was modified: expected ${suggestion.oldText.length} chars, got ${currentText.length}`);
          errors.push(`${suggestion.elementType} #${suggestion.paraIndex} was modified`);
        }
        
        console.log(`⏱️ Operation finished in ${new Date().getTime() - operationStart}ms`);
        
      } catch (e) {
        console.error(`💥 ERROR applying change to ${suggestion.elementType} #${suggestion.paraIndex}:`, e);
        errors.push(`Error in ${suggestion.elementType} #${suggestion.paraIndex}: ${e.message}`);
      }
    });

    // Save the backup for the "Undo" feature.
    try {
      PropertiesService.getDocumentProperties().setProperty('LAST_RUN_BACKUP', JSON.stringify(backupEntries));
      console.log(`💾 Backup of ${backupEntries.length} changes saved for potential undo.`);
    } catch (e) {
      console.warn('⚠️ Could not save undo backup:', e);
    }
    
    // Finalize progress
    PropertiesService.getDocumentProperties().setProperty(PROGRESS_KEY, JSON.stringify({ applied: appliedCount, total: totalSuggestions, done: true }));

    const totalTime = (new Date().getTime() - startTime) / 1000;
    console.log(`\n🏁 Finished applying changes in ${totalTime.toFixed(2)}s.`);
    console.log(`📊 Stats: ${appliedCount} applied, ${errors.length} errors.`);

    let result = `🎉 Applied ${appliedCount}/${approvedSuggestions.length} changes in ${totalTime.toFixed(2)}s.`;
    if (errors.length > 0) {
      result += `\n⚠️ Not applied: ${errors.length}. See logs for details.`;
      console.warn("Error details:", errors);
    }

    return result + `\n↩️ To revert, use the "Undo Last Run" button.`;
    
  } catch (e) {
    console.error("💥💥💥 CRITICAL ERROR in applySuggestions:", e, e.stack);
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
          element.replaceText(escapeRegex(entry.replaceWith), entry.fragment);
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
    return !!key;
  } catch (e) {
    return false;
  }
}

/**
 * Calls OpenAI to find the best semantic match from a list of candidates.
 * @param {string} originalFragment The text to match against.
 * @param {Array<string>} paraArr An array of candidate texts.
 * @returns {string} The index of the best match or "NOT_FOUND".
 */
function callOpenAI_findBestMatchIndex(originalFragment, paraArr) {
  const key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  if (!key) return 'NOT_FOUND';
  if (!paraArr || paraArr.length === 0) return 'NOT_FOUND';

  const systemPrompt = `You are a language expert specializing in semantic analysis. Your task is to find the candidate text that is the closest semantic match to the given fragment.

GUIDELINES:
- Analyze the meaning, not just word-for-word similarity.
- The fragment could be a heading, a summary, or the start of a sentence.
- Respond with ONLY the numeric index (e.g., "0", "1") or the string "NOT_FOUND". Do not add any extra text.`;

  const candidatesText = paraArr.map((p, i) => {
    const preview = p.length > 150 ? p.substring(0, 150) + '...' : p;
    return `[${i}]: ${preview}`;
  }).join('\n\n');
  
  const userPrompt = `FRAGMENT: "${originalFragment}"\n\nCANDIDATES:\n${candidatesText}\n\nIndex of the best match:`;

  const payload = {
    model: 'gpt-4o-mini',
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ],
    temperature: 0.1,
    max_tokens: 10
  };

  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      contentType: 'application/json',  
      headers: { 'Authorization': 'Bearer ' + key },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
     
    const responseCode = response.getResponseCode();
    if (responseCode === 200) {
      const result = JSON.parse(response.getContentText()).choices[0].message.content.trim();
      // Ensure the response is a valid number string or "NOT_FOUND".
      return (!isNaN(parseInt(result, 10)) || result === 'NOT_FOUND') ? result : 'NOT_FOUND';
    } else {
      console.error(`AI API Error: Received HTTP ${responseCode}. Response: ${response.getContentText()}`);
    }
  } catch (e) {
    console.error(`AI call failed: ${e}`);
  }
  
  return 'NOT_FOUND';
} 