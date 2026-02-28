@echo off
TITLE NIPT Report Generator - Setup
cd /d "%~dp0"

echo.
echo ===========================================
echo   NIPT Report Generator - SETUP
echo ===========================================
echo.

REM Check for Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.10+ (64-bit).
    pause
    exit /b 1
)

echo [OK] Python found
echo.

REM Create venv
echo [*] Creating virtual environment...
python -m venv .venv
if errorlevel 1 (
    echo [ERROR] Failed to create virtual environment
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
    echo [ERROR] Failed to install dependencies
    pause
    exit /b 1
)
echo [OK] Setup complete!
echo.
pause
