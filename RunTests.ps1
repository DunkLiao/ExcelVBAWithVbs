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

function Get-TestCounts($output) {
    # pytest summary line 格式：
    #   "729 passed in 1.07s"
    #   "3 failed, 726 passed in 2.06s"
    #   "5 failed, 720 passed, 4 error in 3.12s"
    $sum = $output | Where-Object { $_ -match "\d+ passed" } | Select-Object -Last 1
    $passed = 0; $failed = 0; $errors = 0
    if ($sum -match "(\d+) passed")  { $passed = [int]$Matches[1] }
    if ($sum -match "(\d+) failed")  { $failed = [int]$Matches[1] }
    if ($sum -match "(\d+) error")   { $errors = [int]$Matches[1] }
    $total = $passed + $failed + $errors
    return @{ Total = $total; Passed = $passed; Failed = $failed; Errors = $errors; Line = $sum }
}

function Write-TestStats($label, $counts) {
    $t = $counts.Total; $p = $counts.Passed; $f = $counts.Failed + $counts.Errors
    Write-Host ("  {0,-10} 總計: {1,4}   通過: {2,4}   失敗: {3,4}" -f $label, $t, $p, $f) `
        -ForegroundColor $(if ($f -eq 0) { "Green" } else { "Red" })
}

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
$p1Counts = Get-TestCounts $p1Out

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
Write-TestStats "Phase 1" $p1Counts
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
    $p2Counts = Get-TestCounts $p2Out

    $p2Sum = ($p2Out | Where-Object { $_ -match "passed|failed|error|skipped" } | Select-Object -Last 1)

    if ($p2Exit -ne 0) {
        Write-Host ""
        $p2Out | Where-Object { $_ -match "^FAILED|COM|Error" } | ForEach-Object {
            Write-Host "  $_" -ForegroundColor Red
        }
    }

    Write-Host ""
    Write-TestStats "Phase 2" $p2Counts
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

# 合計所有 Phase 的數字
$p2Total  = if ($IncludeExcel) { $p2Counts.Total }  else { 0 }
$p2Pass   = if ($IncludeExcel) { $p2Counts.Passed } else { 0 }
$p2Fail   = if ($IncludeExcel) { $p2Counts.Failed + $p2Counts.Errors } else { 0 }

$totalAll  = $p1Counts.Total  + $p2Total
$passedAll = $p1Counts.Passed + $p2Pass
$failedAll = ($p1Counts.Failed + $p1Counts.Errors) + $p2Fail

Write-Host ""
Write-Host ("  {0,-12} {1,6}" -f "測試總數：", $totalAll)  -ForegroundColor White
Write-Host ("  {0,-12} {1,6}" -f "通過數：",   $passedAll) -ForegroundColor Green
Write-Host ("  {0,-12} {1,6}" -f "失敗數：",   $failedAll) -ForegroundColor $(if ($failedAll -eq 0) { "Green" } else { "Red" })
Write-Host ""

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