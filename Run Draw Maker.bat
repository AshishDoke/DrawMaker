@echo off
title TT Cybage Draw Maker

echo ============================================================
echo   TT Cybage Internal Draw Maker
echo ============================================================
echo.

:: Check Python is installed
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH.
    echo Please install Python 3.10+ from https://www.python.org/downloads/
    echo Make sure to check "Add Python to PATH" during installation.
    pause
    exit /b 1
)

:: Install dependencies if not present
echo Checking dependencies...
pip show streamlit >nul 2>&1
if errorlevel 1 (
    echo Installing required packages (first-time setup, please wait)...
    pip install -r requirements.txt
    if errorlevel 1 (
        echo ERROR: Failed to install dependencies. Check your internet connection.
        pause
        exit /b 1
    )
)

echo.
echo Starting app... a browser window will open automatically.
echo To stop the app, close this window or press Ctrl+C here.
echo.

streamlit run app.py --server.headless false --browser.gatherUsageStats false

pause
