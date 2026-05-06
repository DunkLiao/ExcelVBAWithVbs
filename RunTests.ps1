<#
.SYNOPSIS
    VBA static test runner. Runs all .bas file checks and shows results.
    Add -IncludeExcel to also run Phase 2 Excel COM compilation.

.EXAMPLE
    .\RunTests.ps1
    .\RunTests.ps1 -IncludeExcel
    .\RunTests.ps1 -Verbose
#>

param(
    [switch]$IncludeExcel,
    [switch]$Verbose
)

function Write-Header($text) {
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "  $text" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
}

function Write-Section($text) {
    Write-Host ""
    Write-Host ">> $text" -ForegroundColor Yellow
    Write-Host ("-" * 50) -ForegroundColor DarkGray
}

function Write-Pass($text) { Write-Host "  [PASS] $text" -ForegroundColor Green }
function Write-Fail($text) { Write-Host "  [FAIL] $text" -ForegroundColor Red }
function Write-Info($text) { Write-Host "  $text" -ForegroundColor Gray }

# ---------------------------------------------------------------------------
Write-Header "VBA Automated Test Runner"
Write-Info "Time : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Write-Info "Dir  : $PSScriptRoot"

# Check Python
$pyVer = & python --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Fail "Python not found. Install Python and add to PATH."
    exit 1
}
$ptVer = & python -m pytest --version 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Fail "pytest not found. Run: pip install pytest"
    exit 1
}
Write-Pass "Python : $pyVer"
Write-Pass "pytest : $ptVer"

# ---------------------------------------------------------------------------
# Phase 1 - Static checks (no Excel required)
# ---------------------------------------------------------------------------
Write-Section "Phase 1 : Static checks (no Excel required)"

$p1Args = if ($Verbose) {
    @("tests/test_static.py", "--tb=short", "-v")
} else {
    @("tests/test_static.py", "--tb=short", "-q", "--no-header")
}

$p1Out  = & python -m pytest @p1Args 2>&1
$p1Exit = $LASTEXITCODE

$p1Sum = ($p1Out | Where-Object { $_ -match "passed|failed|error" } | Select-Object -Last 1)

if ($p1Exit -ne 0) {
    Write-Host ""
    Write-Host "  --- Failures ---" -ForegroundColor DarkRed
    $p1Out | Where-Object { $_ -match "^FAILED" } | ForEach-Object {
        Write-Host "  $_" -ForegroundColor Red
    }
    Write-Host ""
    $p1Out | Where-Object { $_ -match "AssertionError|assert" } | Select-Object -First 20 | ForEach-Object {
        Write-Host "  $_" -ForegroundColor DarkYellow
    }
}

Write-Host ""
if ($p1Exit -eq 0) {
    Write-Pass "Phase 1 ALL PASSED : $p1Sum"
} else {
    Write-Fail "Phase 1 FAILED     : $p1Sum"
}

# ---------------------------------------------------------------------------
# Phase 2 - Excel COM compilation (optional)
# ---------------------------------------------------------------------------
$p2Exit = 0

if ($IncludeExcel) {
    Write-Section "Phase 2 : Excel COM compilation (requires desktop Excel)"
    Write-Info "Prerequisite: Excel > Trust Center > Enable 'Trust access to VBA project object model'"
    Write-Host ""

    $p2Args = if ($Verbose) {
        @("tests/test_compile.py", "--tb=short", "-v")
    } else {
        @("tests/test_compile.py", "--tb=short", "-q", "--no-header")
    }

    $p2Out  = & python -m pytest @p2Args 2>&1
    $p2Exit = $LASTEXITCODE

    $p2Sum = ($p2Out | Where-Object { $_ -match "passed|failed|error|skipped" } | Select-Object -Last 1)

    if ($p2Exit -ne 0) {
        Write-Host ""
        $p2Out | Where-Object { $_ -match "^FAILED|COM|Error" } | ForEach-Object {
            Write-Host "  $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    if ($p2Exit -eq 0) {
        Write-Pass "Phase 2 ALL PASSED : $p2Sum"
    } else {
        Write-Fail "Phase 2 FAILED     : $p2Sum"
    }
} else {
    Write-Info "(Phase 2 skipped. Add -IncludeExcel to enable Excel COM compilation test.)"
}

# ---------------------------------------------------------------------------
# Summary
# ---------------------------------------------------------------------------
Write-Header "Test Summary"

$overall = if ($p1Exit -ne 0 -or $p2Exit -ne 0) { 1 } else { 0 }

if ($overall -eq 0) {
    Write-Host "  [OK] All tests passed!" -ForegroundColor Green
} else {
    Write-Host "  [FAIL] Some tests failed. See details above." -ForegroundColor Red
    Write-Host ""
    Write-Host "  To see full detail:" -ForegroundColor Yellow
    Write-Host "    python -m pytest tests/test_static.py -v" -ForegroundColor Gray
}

Write-Host ""
exit $overall