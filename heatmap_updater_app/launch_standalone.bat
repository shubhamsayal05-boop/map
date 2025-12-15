@echo off
REM Launcher for Standalone HeatMap Application (Windows)

title AVLDrive HeatMap Manager - Standalone
echo ============================================================
echo AVLDrive HeatMap Manager - Standalone Application
echo ============================================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7+ from https://www.python.org
    pause
    exit /b 1
)

echo Starting standalone application...
echo (No external dependencies required)
echo.

REM Launch the standalone application
python launch_standalone.py

if errorlevel 1 (
    echo.
    echo ERROR: Application exited with an error
    pause
)
