@echo off
title Invoice Generator — Launcher
echo =============================================
echo   Invoice Generator by Barry's Gamebox
echo =============================================
echo.

:: Check Python
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python not found. Please install Python 3.9+ from https://python.org
    pause
    exit /b 1
)

echo [1/3] Checking dependencies...
python -m pip install Pillow reportlab --quiet --upgrade

echo [2/3] Starting Invoice Generator...
echo.
python "%~dp0invoice_app.py"

if errorlevel 1 (
    echo.
    echo [ERROR] Application crashed. Check the error above.
    pause
)
