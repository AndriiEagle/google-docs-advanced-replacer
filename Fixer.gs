/* ───────── 🤖 Smart Fragment Fixer ───────── */

/**
 * Интеллектуальная система исправления мелких проблем во фрагментах.
 * Автоматически находит и исправляет проблемы с символами, пробелами, кавычками.
 * @param {Array<Object>} directives Исходные директивы.
 * @param {Array<Object>} allElements Все элементы документа.
 * @param {boolean} hasApiKey Наличие OpenAI API ключа.
 * @returns {Array<Object>} Исправленные директивы.
 */
function applySmartFragmentFixer(directives, allElements, hasApiKey) {
  logOperation('INFO', '🤖 Smart Fragment Fixer: начинаю анализ директив');
  
  return directives.map((directive, index) => {
    const originalFragment = directive.fragment;
    if (!originalFragment) {
      // Пропускаем директивы без фрагмента
      return { ...directive, wasFixed: false, fixType: [] };
    }
    
    let currentFragment = originalFragment;
    const fixType = [];
    
    // STEP 1: Check if basic normalization is enough
    const hasNormalizedMatch = allElements.some(elem => 
      normalizeText(elem.text).includes(normalizeText(originalFragment))
    );
    
    if (hasNormalizedMatch) {
      // Normalization was sufficient, no manual fixes needed
      return { ...directive, fragment: originalFragment, wasFixed: false, fixType: [] };
    }
    
    // STEP 2: Apply manual fixes only if normalization failed
    logOperation('WARN', `⚠️ Директива #${index + 1}: нормализация не помогла, применяю ручные исправления`);

    // Этапы исправления
    const fixes = [
      (frag) => frag.replace(/\s-\s/g, ' — ').replace(/^-\s/g, '— ').replace(/\s-$/g, ' —'), // Тире
      (frag) => frag.replace(/\"(.*?)\"/g, '«$1»'), // Кавычки
      (frag) => frag.replace(/\s+/g, ' ').trim(), // Пробелы
      (frag) => frag.replace(/[\u2019'`]/g, 'ʼ') // Апострофы (legacy)
    ];

    const fixNames = ['тире', 'кавычки', 'пробелы', 'апострофы'];

    fixes.forEach((fix, i) => {
      const fixed = fix(currentFragment);
      if (fixed !== currentFragment) {
        fixType.push(fixNames[i]);
        currentFragment = fixed;
      }
    });

    const wasFixed = fixType.length > 0;
    if (wasFixed) {
       logOperation('INFO', `🔧 Директива #${index + 1}: применены исправления (${fixType.join(', ')}). "${originalFragment}" -> "${currentFragment}"`);
    }
    
    return { ...directive, fragment: currentFragment, originalFragment, wasFixed, fixType };
  });
}
