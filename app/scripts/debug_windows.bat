@echo off
title Windows Debug - WinDaq Analyzer
color 0A

echo.
echo ========================================
echo    Windows Debug - WinDaq Analyzer
echo ========================================
echo.

echo [DEBUG] Testing Python installations...
echo.

echo Testing 'python' command:
python --version
if %errorlevel% equ 0 (
    echo [OK] 'python' command works
    echo Testing Python packages with 'python':
    python -c "import pandas, openpyxl, numpy, scipy, matplotlib; print('[OK] All packages available')"
    if %errorlevel% equ 0 (
        echo [OK] All Python packages work with 'python'
    ) else (
        echo [ERROR] Some packages missing with 'python'
    )
) else (
    echo [ERROR] 'python' command not found
)

echo.
echo Testing 'python3' command:
python3 --version
if %errorlevel% equ 0 (
    echo [OK] 'python3' command works
    echo Testing Python packages with 'python3':
    python3 -c "import pandas, openpyxl, numpy, scipy, matplotlib; print('[OK] All packages available')"
    if %errorlevel% equ 0 (
        echo [OK] All Python packages work with 'python3'
    ) else (
        echo [ERROR] Some packages missing with 'python3'
    )
) else (
    echo [ERROR] 'python3' command not found
)

echo.
echo Testing 'py' command:
py --version
if %errorlevel% equ 0 (
    echo [OK] 'py' command works
    echo Testing Python packages with 'py':
    py -c "import pandas, openpyxl, numpy, scipy, matplotlib; print('[OK] All packages available')"
    if %errorlevel% equ 0 (
        echo [OK] All Python packages work with 'py'
    ) else (
        echo [ERROR] Some packages missing with 'py'
    )
) else (
    echo [ERROR] 'py' command not found
)

echo.
echo [DEBUG] Testing Node.js...
node --version
if %errorlevel% equ 0 (
    echo [OK] Node.js is working
) else (
    echo [ERROR] Node.js not found
)

echo.
echo [DEBUG] Testing npm...
npm --version
if %errorlevel% equ 0 (
    echo [OK] npm is working
) else (
    echo [ERROR] npm not found
)

echo.
echo [DEBUG] Current directory contents:
dir *.py
echo.
dir *.js
echo.

echo [DEBUG] Environment PATH:
echo %PATH%

echo.
echo ========================================
echo    Debug Complete
echo ========================================
echo.
pause
