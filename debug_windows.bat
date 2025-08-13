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
    echo ✓ 'python' command works
    echo Testing Python packages with 'python':
    python -c "import pandas, openpyxl, numpy, scipy, matplotlib; print('✓ All packages available')"
    if %errorlevel% equ 0 (
        echo ✓ All Python packages work with 'python'
    ) else (
        echo ❌ Some packages missing with 'python'
    )
) else (
    echo ❌ 'python' command not found
)

echo.
echo Testing 'python3' command:
python3 --version
if %errorlevel% equ 0 (
    echo ✓ 'python3' command works
    echo Testing Python packages with 'python3':
    python3 -c "import pandas, openpyxl, numpy, scipy, matplotlib; print('✓ All packages available')"
    if %errorlevel% equ 0 (
        echo ✓ All Python packages work with 'python3'
    ) else (
        echo ❌ Some packages missing with 'python3'
    )
) else (
    echo ❌ 'python3' command not found
)

echo.
echo Testing 'py' command:
py --version
if %errorlevel% equ 0 (
    echo ✓ 'py' command works
    echo Testing Python packages with 'py':
    py -c "import pandas, openpyxl, numpy, scipy, matplotlib; print('✓ All packages available')"
    if %errorlevel% equ 0 (
        echo ✓ All Python packages work with 'py'
    ) else (
        echo ❌ Some packages missing with 'py'
    )
) else (
    echo ❌ 'py' command not found
)

echo.
echo [DEBUG] Testing Node.js...
node --version
if %errorlevel% equ 0 (
    echo ✓ Node.js is working
) else (
    echo ❌ Node.js not found
)

echo.
echo [DEBUG] Testing npm...
npm --version
if %errorlevel% equ 0 (
    echo ✓ npm is working
) else (
    echo ❌ npm not found
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
