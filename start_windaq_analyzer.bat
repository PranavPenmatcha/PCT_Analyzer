@echo off
chcp 65001 >nul 2>&1
set PYTHONIOENCODING=utf-8
set PYTHONLEGACYWINDOWSSTDIO=1
title WinDaq Analyzer Startup
color 0A

echo.
echo ========================================
echo    WinDaq Analyzer - Starting Up
echo ========================================
echo.

REM Check if Node.js is installed
echo [1/5] Checking Node.js installation...
node --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Node.js is not installed!
    echo Please install Node.js from https://nodejs.org/
    echo.
    pause
    exit /b 1
)
echo [OK] Node.js is installed

REM Check if Python is installed
echo [2/5] Checking Python installation...
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo ERROR: Python is not installed!
    echo Please install Python from https://python.org/
    echo.
    pause
    exit /b 1
)
echo [OK] Python is installed

REM Install Node.js dependencies
echo [3/5] Installing Node.js dependencies...
if not exist "node_modules" (
    echo Installing npm packages...
    call npm install
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install npm packages!
        pause
        exit /b 1
    )
) else (
    echo [OK] Node.js dependencies already installed
)

REM Install Python dependencies
echo [4/5] Installing Python dependencies...
echo Installing Python packages...
pip install pandas openpyxl numpy scipy matplotlib xlsxwriter
if %errorlevel% neq 0 (
    echo WARNING: Some Python packages may have failed to install
    echo Trying with python -m pip...
    python -m pip install pandas openpyxl numpy scipy matplotlib xlsxwriter
    if %errorlevel% neq 0 (
        echo ERROR: Failed to install Python packages!
        echo Please install manually: pip install pandas openpyxl numpy scipy matplotlib xlsxwriter
        pause
    )
)
echo [OK] Python dependencies installed

REM Start the application
echo [5/5] Starting WinDaq Analyzer...
echo.
echo ========================================
echo   WinDaq Analyzer is starting...
echo   Opening browser in 3 seconds...
echo ========================================
echo.

REM Start the server in background and open browser
start /B npm start
timeout /t 3 /nobreak >nul

REM Open browser
start http://localhost:3000

echo.
echo [OK] WinDaq Analyzer is now running!
echo [OK] Browser should open automatically
echo.
echo Instructions:
echo - Upload your WinDaq files (.wdq, .wdh, .wdc)
echo - View automatic pulse analysis
echo - Download Excel files with charts
echo.
echo To stop the server, close this window or press Ctrl+C
echo.

REM Keep the window open to show server logs
echo Press Ctrl+C to stop the server
echo.
npm start
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Server failed to start!
    echo Check the error messages above
    pause
)
