@echo off
setlocal EnableExtensions
title IBKR Trade Analysis
cd /d "%~dp0"

set "PY_CMD="
where py >nul 2>&1 && set "PY_CMD=py -3"
if not defined PY_CMD (
    where python >nul 2>&1 && set "PY_CMD=python"
)
if not defined PY_CMD (
    echo Python 3 was not found on PATH.
    echo Install Python from https://www.python.org/downloads/ and check "Add python.exe to PATH".
    pause
    exit /b 1
)

%PY_CMD% -c "import streamlit,pandas,plotly,openpyxl" >nul 2>&1
if errorlevel 1 (
    echo Installing required packages...
    %PY_CMD% -m pip install -r "%~dp0requirements.txt"
    if errorlevel 1 (
        echo.
        echo Could not install dependencies. Check your internet connection and try again.
        pause
        exit /b 1
    )
)

echo Starting IBKR Trade Analysis...
echo Streamlit will open the app in your default browser.
echo Close this window, or press Ctrl+C, to stop the server.
echo.

REM headless=false tells Streamlit to launch the default browser once the server is ready.
%PY_CMD% -m streamlit run "%~dp0app.py" --server.headless=false --browser.gatherUsageStats=false
set "RC=%ERRORLEVEL%"
if not "%RC%"=="0" (
    echo.
    echo The app stopped with exit code %RC%.
    pause
)
endlocal & exit /b %RC%
