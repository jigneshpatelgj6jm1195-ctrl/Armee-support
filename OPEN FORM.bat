@echo off
title School Complaint Form Launcher
echo.
echo  =============================================
echo    School Equipment Complaint Form Launcher
echo  =============================================
echo.
echo  Starting... Please wait.
echo.

:: Try python first, then py launcher
python "%~dp0export_and_launch.py"
if %errorlevel% neq 0 (
    py "%~dp0export_and_launch.py"
)
if %errorlevel% neq 0 (
    echo.
    echo ERROR: Python not found or script failed!
    echo Please install Python from https://python.org
    echo.
    pause
)
pause
