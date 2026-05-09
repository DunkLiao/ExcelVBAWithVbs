@echo off
echo ================================================
echo  Excel VBA Module Library - Build Tool
echo ================================================
echo.

cd /d "%~dp0.."

python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python and add it to PATH.
    echo.
    pause
    exit /b 1
)

echo Running build script...
echo.
python frontend\build.py
echo.

if errorlevel 1 (
    echo [ERROR] Build failed. See messages above.
) else (
    echo Build complete! Open frontend\index.html in your browser.
)

echo.
pause