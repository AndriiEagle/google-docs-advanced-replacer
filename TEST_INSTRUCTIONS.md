# Instructions for running Advanced Replacer tests

## HOW TO RUN THE TEST

1. Open Google Apps Script (script.google.com).
2. Create a new project or open an existing one.
3. Copy the content of all `.gs` files (`Core.gs`, `Document.gs`, `Utils.gs`, `Logging.gs`, `Fixer.gs`, `AISystem.gs`, `TestRunner.gs`) into the project, creating new script files for each.
4. Copy the content of all `.html` files (`JavaScript.html`, `Sidebar.html`) into the project, creating new HTML files for each.
5. In the script editor, from the function dropdown, select the function to run:
   - `runUserDirectivesTest()` - to test the 15 user directives.
   - `runAllTests()` - for a full test suite.
   - `runSimpleTest()` - for a basic functionality test.
6. Click the "Run" button (▶️).
7. Check the results in the execution log (View > Logs or `Ctrl+Enter`).

## OpenAI API SETUP (Optional, for AI features)

**Method 1: Through UI (Recommended)**
1. Get an API key from openai.com.
2. Open Advanced Replacer in Google Docs.
3. Click the 🔑 button in the top right corner.
4. Enter your API key and click "💾 Сохранить".
5. Test with "🧪 Тест" button.

**Method 2: Through Google Apps Script (Alternative)**
1. Get an API key from openai.com.
2. In your Google Apps Script project, go to "Project Settings" (⚙️ icon).
3. Under "Script Properties", click "Add script property".
4. Add a new property with the following details:
   - **Property name:** `OPENAI_API_KEY`
   - **Value:** `your_key_here` (paste your actual OpenAI API key)

**Note:** The system will automatically detect API keys from both locations for compatibility.

## EXPECTED RESULTS

- **Matches found:** 15/15 (100% success rate).
- **Match types:** 15 "EXACT" matches.
- **AI matches:** 0 (since all should be found exactly).
- **Smart Fixer:** 0 to 5 corrections are considered normal as the fixer adjusts minor text inconsistencies.
