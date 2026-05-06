@echo off
chcp 65001 > nul
cd /d "%~dp0"

echo.
echo ============================================================
echo  VBA 自動化測試執行器
echo ============================================================
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0RunTests.ps1" %*

if %ERRORLEVEL% neq 0 (
    echo.
    echo  [失敗] 測試未全部通過。
    echo.
    pause
    exit /b 1
) else (
    echo.
    echo  [完成] 測試全部通過。
    echo.
    pause
    exit /b 0
)
