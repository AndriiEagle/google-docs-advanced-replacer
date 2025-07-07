/* ───────── 📄 Document Processing ───────── */

/**
 * Safely gets the body of the active document.
 * @returns {GoogleAppsScript.Document.Body} The document body.
 * @throws {Error} If no active document or body is found.
 */
function safeGetBody() {
  logOperation('INFO', '📄 Getting document body...');
  const doc = DocumentApp.getActiveDocument();
  if (!doc) {
    logOperation('ERROR', 'No active Google Docs document found');
    throw new Error('No active Google Docs document found.');
  }
  const body = doc.getBody();
  if (!body) {
    logOperation('ERROR', 'Could not retrieve the document body');
    throw new Error('Could not retrieve the document body.');
  }
  logOperation('INFO', '✅ Document body retrieved successfully.');
  return body;
}

/**
 * Retrieves all processable text-containing elements from the document.
 * This function recursively traverses the document structure.
 * @returns {Array<Object>} An array of element wrappers with unique IDs.
 */
function getAllDocumentElements() {
  logOperation('INFO', '📄 Starting document elements retrieval...');
  const startTime = new Date().getTime();
  const body = safeGetBody();
  const elements = [];
  let index = 0;
  const elementTypeCounts = {};

  // Define element types that we consider as primary, self-contained text blocks.
  const leafTypes = [
    DocumentApp.ElementType.PARAGRAPH,
    DocumentApp.ElementType.LIST_ITEM,
    DocumentApp.ElementType.TABLE_CELL
  ];

  function processContainer(container) {
    if (!container || typeof container.getNumChildren !== 'function') return;

    for (let i = 0; i < container.getNumChildren(); i++) {
      const element = container.getChild(i);
      const elementType = element.getType();

      // If the element is a "leaf" type, process it and stop recursion.
      if (leafTypes.includes(elementType)) {
        const text = getElementText(element);
        if (text && text.trim().length > 0) {
          const elementTypeName = getElementTypeName(elementType);
          elementTypeCounts[elementTypeName] = (elementTypeCounts[elementTypeName] || 0) + 1;
          const elementWrapper = {
            element: element,
            text: text.trim(),
            type: elementType,
            typeName: elementTypeName,
            originalIndex: index++
          };
          elementWrapper.id = generateElementId(elementWrapper);
          elements.push(elementWrapper);
        }
      } else if (typeof element.getNumChildren === 'function' && element.getNumChildren() > 0) {
        // If it's another type of container (like a Table), recurse into it.
        processContainer(element);
      }
    }
  }

  processContainer(body);
  
  const processingTime = new Date().getTime() - startTime;
  logOperation('INFO', `✅ Document parsing complete: Found ${elements.length} elements in ${processingTime}ms`, {
    totalElements: elements.length,
    processingTime: processingTime,
    elementTypeCounts: elementTypeCounts
  });
  
  return elements;
}

/**
 * Intelligently extracts text from any Google Docs element type.
 * @param {GoogleAppsScript.Document.Element} element The document element.
 * @returns {string} The extracted text, or an empty string if not applicable.
 */
function getElementText(element) {
  try {
    const elementType = element.getType();
    if (element.getText && typeof element.getText === 'function') {
      return element.getText();
    }
    // Fallback for elements that might not have getText but are text-based.
    if (element.asText && typeof element.asText === 'function') {
      return element.asText().getText();
    }
  } catch (e) {
    // Ignore elements we can't get text from.
  }
  return '';
}

/**
 * A utility to apply text changes to various element types.
 * @param {GoogleAppsScript.Document.Element} element The element to modify.
 * @param {string} newText The new text to apply.
 */
function applyTextChange(element, newText) {
  const elementType = element.getType();
  try {
    // Universal approach for text-based elements that support it.
    if (element.clear && typeof element.clear === 'function' && element.setText && typeof element.setText === 'function') {
        element.clear().setText(newText);
    } else if (element.setText && typeof element.setText === 'function') {
        element.setText(newText);
    } else {
        throw new Error(`Element type ${elementType} does not support text modification.`);
    }
  } catch (e) {
    logOperation('ERROR', `💥 Failed to modify text for element type ${elementType}: ${e.message}`, { error: e });
    throw e; // Re-throw to be caught by the caller
  }
}
