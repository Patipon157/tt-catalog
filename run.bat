@echo off
title Thailand Trophy - Program Launcher
echo ============================================
echo   Thailand Trophy - Program Launcher
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

echo  เลือกโปรแกรมที่ต้องการเปิด:
echo.
echo  [1] Catalog Manager    - จัดการไฟล์ Catalog
echo  [2] Product Master     - ต้นทุนสินค้า (Master File)
echo  [3] Check Folder       - ตรวจสอบการเข้าถึง Folder
echo  [0] ออก
echo.
set /p choice="  เลือก (1/2/3/0): "

if "%choice%"=="1" (
    echo.
    echo  Starting Catalog Manager...
    python "%~dp0catalog_manager.py"
    goto :end
)

if "%choice%"=="2" (
    echo.
    echo  Starting Product Master...
    python "%~dp0product_master.py"
    goto :end
)

if "%choice%"=="3" (
    echo.
    echo  Checking folder access...
    python "%~dp0check_folder_access.py"
    echo.
    pause
    goto :end
)

if "%choice%"=="0" (
    exit /b 0
)

echo.
echo  [ERROR] Invalid choice. Please enter 1, 2, 3, or 0.
pause

:end
if %ERRORLEVEL% neq 0 (
    echo.
    echo [ERROR] Program exited with an error.
    pause
)
