@echo off
TITLE NIPT Report Generator - Setup
cd /d "%~dp0"

echo.
echo ===========================================
echo   NIPT Report Generator - SETUP
echo ===========================================
echo.

REM Check for Python (try python then python3)
SET PY_CMD=python
%PY_CMD% --version >nul 2>&1
if errorlevel 1 (
    SET PY_CMD=python3
    %PY_CMD% --version >nul 2>&1
    if errorlevel 1 (
        echo [ERROR] Python not found. 
        echo Please ensure Python 3.10+ is installed and 'Add to PATH' was checked.
        pause
        exit /b 1
    )
)

echo [OK] Found: %PY_CMD%
%PY_CMD% --version
echo.

REM Create venv
echo [*] Creating virtual environment...
%PY_CMD% -m venv .venv
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment. 
    echo This can happen if Python was installed from the Windows Store without proper permissions.
    pause
    exit /b 1
)
echo [OK] Created
echo.

REM Activate and install
echo [*] Installing dependencies...
call .venv\Scripts\activate.bat
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
if errorlevel 1 (
    echo.
    echo [ERROR] Failed to install dependencies. 
    echo Please check your internet connection or requirements.txt.
    pause
    exit /b 1
)

echo.
echo [OK] Setup complete!
echo You can now use launch_nipt.bat to run the software.
echo.
pause
