/* ───────── 🤖 AI System ───────── */

/**
 * 🚀 SUPER-OPTIMIZED AI SEARCH SYSTEM
 * Handles large documents efficiently with smart filtering and batching.
 */

/**
 * Calculates a basic similarity score between two strings.
 * @param {string} str1 The first string.
 * @param {string} str2 The second string.
 * @returns {number} A similarity score between 0 and 1.
 */
function calculateBasicSimilarity(str1, str2) {
    if (!str1 || !str2) return 0;
    const words1 = new Set(normalizeText(str1).split(/\s+/));
    const words2 = new Set(normalizeText(str2).split(/\s+/));
    const intersection = new Set([...words1].filter(word => words2.has(word)));
    const union = new Set([...words1, ...words2]);
    return union.size === 0 ? 0 : intersection.size / union.size;
}

/**
 * Multi-stage AI candidate filtering for large documents.
 * @param {string} fragment The search fragment.
 * @param {Array<Object>} allElements All document elements.
 * @param {number} maxCandidates Maximum candidates to return.
 * @returns {Array<Object>} Filtered and scored candidates.
 */
function performAdvancedAIFiltering(fragment, allElements, maxCandidates = 3) {
  return allElements
    .map(elem => ({
      element: elem,
      similarity: calculateBasicSimilarity(fragment, elem.text)
    }))
    .filter(c => c.similarity > 0.2) // Increased threshold for better candidates
    .sort((a, b) => b.similarity - a.similarity)
    .slice(0, maxCandidates);
}

/**
 * Optimized batch AI processing.
 * @param {Array<Object>} aiRequests Array of AI requests to process.
 * @returns {Array<Object>} Results for each request.
 */
function processBatchAIRequests(aiRequests) {
  // This is a placeholder for a more complex batching system.
  // For now, it processes requests one by one.
  return aiRequests.map(request => {
    const cacheKey = `${request.fragment}_${request.candidates.map(c => c.element.id).join('|')}`;
    if (AI_SEARCH_CACHE.has(cacheKey)) {
      logOperation('INFO', `💾 AI Cache hit for fragment: "${request.fragment.substring(0,30)}..."`);
      return AI_SEARCH_CACHE.get(cacheKey);
    }

    if (request.candidates.length > 0) {
      const candidateTexts = request.candidates.map(c => c.element.text);
      const aiResult = callOptimizedOpenAI_findBestMatch(request.fragment, candidateTexts);
      
      let bestMatch = null;
      if (aiResult !== 'NOT_FOUND' && !isNaN(parseInt(aiResult, 10))) {
        const index = parseInt(aiResult, 10);
        if (index >= 0 && index < request.candidates.length) {
          bestMatch = request.candidates[index];
        }
      }
      const result = { fragment: request.fragment, bestMatch, aiResult };
      AI_SEARCH_CACHE.set(cacheKey, result);
      return result;
    }
    return { fragment: request.fragment, bestMatch: null, aiResult: 'NO_CANDIDATES' };
  });
}

/**
 * Optimized OpenAI call with better token management.
 * @param {string} fragment The fragment to search for.
 * @param {Array<string>} candidateTexts Array of candidate texts.
 * @returns {string} The index of the best match or "NOT_FOUND".
 */
function callOptimizedOpenAI_findBestMatch(fragment, candidateTexts) {
  // Сначала проверяем Document Properties (новый способ)
  let key = PropertiesService.getDocumentProperties().getProperty('OPENAI_API_KEY');
  
  // Если не найден, проверяем Script Properties (старый способ)
  if (!key) {
    key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  }
  
  if (!key) return 'NOT_FOUND';

  const systemPrompt = `You are \"SigmaText Fragment Selector\" – a multilingual semantic matcher.
Given QUERY_FRAGMENT and a numbered CANDIDATE list, pick the index of the candidate whose MEANING best matches the fragment, tolerating minor spelling, punctuation and look-alike letter differences. If none reaches sufficient similarity, output NOT_FOUND.

Respond ONLY with the numeric index (0,1,2,…) or NOT_FOUND.`;
  const userPrompt = `FRAGMENT: "${fragment}"\n\nCANDIDATES:\n${candidateTexts.map((text, i) => `[${i}]: ${text}`).join('\n')}`;

  const payload = {
    model: 'gpt-4o-mini',
    messages: [{ role: 'system', content: systemPrompt }, { role: 'user', content: userPrompt }],
    temperature: 0.1,
    max_tokens: 5
  };

  return callOpenAI_findBestMatchIndex_Internal(payload, candidateTexts.length);
}

/**
 * Internal AI call with retry logic.
 * @param {Object} payload The API payload.
 * @param {number} maxIndex Maximum valid index.
 * @returns {string} Result or "NOT_FOUND".
 */
function callOpenAI_findBestMatchIndex_Internal(payload, maxIndex) {
  // Сначала проверяем Document Properties (новый способ)
  let key = PropertiesService.getDocumentProperties().getProperty('OPENAI_API_KEY');
  
  // Если не найден, проверяем Script Properties (старый способ)
  if (!key) {
    key = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
  }
  
  const MAX_RETRIES = 3;
  const BASE_DELAY = 1000;
  
  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    try {
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
        const result = JSON.parse(responseText).choices[0].message.content.trim();
        const index = parseInt(result, 10);
        if (!isNaN(index) && index >= 0 && index < maxIndex) return result;
        return 'NOT_FOUND';
      }
      Utilities.sleep(BASE_DELAY * Math.pow(2, attempt));
    } catch (e) {
      if (attempt >= MAX_RETRIES - 1) {
        logOperation('ERROR', `💥 AI API call failed after ${MAX_RETRIES} attempts: ${e.message}`);
      }
    }
  }
  return 'NOT_FOUND';
}

/**
 * Clears the AI search cache.
 */
function clearAISearchCache() {
  AI_SEARCH_CACHE.clear();
  logOperation('INFO', '🗑️ AI search cache cleared');
}

/**
 * Manages AI cache size to prevent memory issues.
 */
function manageAICacheSize() {
  const maxSize = 100;
  if (AI_SEARCH_CACHE.size > maxSize) {
    const entries = Array.from(AI_SEARCH_CACHE.keys());
    for (let i = 0; i < AI_SEARCH_CACHE.size - maxSize; i++) {
      AI_SEARCH_CACHE.delete(entries[i]);
    }
    logOperation('INFO', `🧹 AI cache trimmed to ${maxSize} entries`);
  }
}
