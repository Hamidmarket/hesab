@echo off
chcp 65001 >nul
title Hamid Market - Customer Update
cd /d "%~dp0"

echo ========================================
echo       Hamid Market - Customer Update
echo ========================================
echo.
echo Starting update.py ...
echo.

python update.py

if errorlevel 1 (
    echo.
    echo ERROR: Update failed!
) else (
    echo.
    echo SUCCESS: Update completed!
)

echo.
pause