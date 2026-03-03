@echo off
title Thailand Trophy - Catalog Manager
echo ============================================
echo   Thailand Trophy - Catalog Manager
echo ============================================
echo.

:: Check if Python is installed
python --version >nul 2>&1
if %ERRORLEVEL% neq 0 (
    echo [ERROR] Python not found!
    echo.
    echo Please install Python from:
    echo   https://www.python.org/downloads/
    echo.
    echo When installing, check "Add Python to PATH"
    echo.
    pause
    exit /b 1
)

echo Starting Catalog Manager...
echo.
python "%~dp0catalog_manager.py"

if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Program exited with an error.
    pause
)
