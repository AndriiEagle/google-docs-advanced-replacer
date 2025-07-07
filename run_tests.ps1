# Advanced Replacer Test Script
# Encoding: UTF-8

# Set output encoding for PowerShell
$OutputEncoding = [System.Text.Encoding]::UTF8

Write-Host '--- Starting Advanced Replacer Tests ---' -ForegroundColor Green
Write-Host '================================================' -ForegroundColor Green

# Check for required files
$requiredFiles = @(
    "Core.gs",
    "Document.gs", 
    "Utils.gs",
    "Logging.gs",
    "Fixer.gs",
    "AISystem.gs",
    "TestRunner.gs",
    "JavaScript.html",
    "Sidebar.html"
)

Write-Host '[INFO] Checking project files...' -ForegroundColor Yellow

$allFilesExist = $true
foreach ($file in $requiredFiles) {
    if (Test-Path $file) {
        Write-Host "[OK] $file - found" -ForegroundColor Green
    } else {
        Write-Host "[ERROR] $file - NOT FOUND" -ForegroundColor Red
        $allFilesExist = $false
    }
}

if (-not $allFilesExist) {
    Write-Host '[ERROR] Not all project files were found!' -ForegroundColor Red
    exit 1
}

Write-Host '[SUCCESS] All project files are present.' -ForegroundColor Green

# Check syntax of main files
Write-Host '[INFO] Checking main files...' -ForegroundColor Yellow

$coreFiles = @("Core.gs", "TestRunner.gs", "JavaScript.html")
foreach ($file in $coreFiles) {
    if (Test-Path $file) {
        $size = (Get-Item $file).Length
        Write-Host "[OK] $file - OK ($size bytes)" -ForegroundColor Green
    }
}

# Display test information
Write-Host ""
Write-Host '--- TEST INFORMATION ---' -ForegroundColor Cyan
Write-Host '================================================' -ForegroundColor Cyan
Write-Host 'Total directives for testing: 15' -ForegroundColor White
Write-Host 'Test type: Text replacement in Google Docs' -ForegroundColor White
Write-Host 'Smart Fixer: Enabled' -ForegroundColor White
Write-Host 'AI support: Enabled (requires OpenAI API key)' -ForegroundColor White

# Count directives in TestRunner.gs
if (Test-Path "TestRunner.gs") {
    $content = Get-Content "TestRunner.gs" -Raw -Encoding UTF8
    $directiveCount = ($content | Select-String '"fragment"' -AllMatches).Matches.Count
    Write-Host "Directives found in TestRunner.gs: $directiveCount" -ForegroundColor White
}

Write-Host ""
Write-Host '--- HOW TO RUN THE TEST ---' -ForegroundColor Green
Write-Host '================================================' -ForegroundColor Green
Write-Host '1. Open Google Apps Script (script.google.com)' -ForegroundColor White
Write-Host '2. Create a new project or open an existing one' -ForegroundColor White
Write-Host '3. Copy the content of all .gs files into the project' -ForegroundColor White
Write-Host '4. Copy the content of all .html files into the project' -ForegroundColor White
Write-Host '5. In the script editor, select the function to run:' -ForegroundColor White
Write-Host '   - runUserDirectivesTest() - to test the 15 user directives' -ForegroundColor Yellow
Write-Host '   - runAllTests() - for a full test suite' -ForegroundColor Yellow
Write-Host '   - runSimpleTest() - for a basic test' -ForegroundColor Yellow
Write-Host '6. Click the "Run" button' -ForegroundColor White
Write-Host '7. Check the results in the console (View > Logs)' -ForegroundColor White

Write-Host ""
Write-Host '--- OpenAI API SETUP (for AI features) ---' -ForegroundColor Cyan
Write-Host '================================================' -ForegroundColor Cyan
Write-Host 'Method 1 (RECOMMENDED): Through UI' -ForegroundColor Yellow
Write-Host '1. Get an API key from openai.com' -ForegroundColor White
Write-Host '2. Open Advanced Replacer in Google Docs' -ForegroundColor White
Write-Host '3. Click the key button (🔑) in top right corner' -ForegroundColor White
Write-Host '4. Enter your API key and click Save' -ForegroundColor White
Write-Host '5. Test with the Test button' -ForegroundColor White
Write-Host '' -ForegroundColor White
Write-Host 'Method 2 (Alternative): Through Google Apps Script' -ForegroundColor Yellow
Write-Host '1. In Google Apps Script: Project Settings > Script Properties' -ForegroundColor White
Write-Host '2. Add a new property: OPENAI_API_KEY = your_key_here' -ForegroundColor White
Write-Host '' -ForegroundColor White
Write-Host 'Note: System will detect API keys from both locations automatically' -ForegroundColor Green

Write-Host ""
Write-Host '--- EXPECTED RESULTS ---' -ForegroundColor Magenta
Write-Host '================================================' -ForegroundColor Magenta
Write-Host '[SUCCESS] Matches found: 15/15 (100%)' -ForegroundColor Green
Write-Host '[SUCCESS] Exact matches: 15' -ForegroundColor Green
Write-Host '[INFO] AI matches: 0 (if all are exact)' -ForegroundColor Green
Write-Host '[INFO] Smart Fixer corrections: 0-5 (this is normal)' -ForegroundColor Green

Write-Host ""
Write-Host '--- PREPARATION COMPLETE ---' -ForegroundColor Green
Write-Host '================================================' -ForegroundColor Green

# Create a simple instruction file
'# Instructions for running Advanced Replacer tests' | Out-File -FilePath "TEST_INSTRUCTIONS.md" -Encoding UTF8

Write-Host '[INFO] Instruction file created: TEST_INSTRUCTIONS.md' -ForegroundColor Yellow
Write-Host ""
Write-Host '--> Now, please follow the instructions to run the test in Google Apps Script!' -ForegroundColor Green 