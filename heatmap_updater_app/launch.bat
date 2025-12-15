@echo off
REM Launch script for HeatMap Updater GUI (Windows)

title AVLDrive HeatMap Updater
echo ===========================================
echo AVLDrive HeatMap Updater - Starting...
echo ===========================================
echo.

REM Change to script directory
cd /d "%~dp0"

REM Check if Python is available
python --version >nul 2>&1
if errorlevel 1 (
    echo ERROR: Python is not installed or not in PATH
    echo Please install Python 3.7 or higher from https://www.python.org
    echo Make sure to check "Add Python to PATH" during installation
    pause
    exit /b 1
)

REM Display Python version
echo Python version:
python --version
echo.

REM Check if openpyxl is installed
python -c "import openpyxl" >nul 2>&1
if errorlevel 1 (
    echo WARNING: openpyxl is not installed
    
    REM Check if pip is available
    pip --version >nul 2>&1
    if errorlevel 1 (
        echo ERROR: pip is not installed
        echo Please install pip or install openpyxl manually
        pause
        exit /b 1
    )
    
    echo Attempting to install dependencies...
    echo.
    pip install -r requirements.txt
    
    if errorlevel 1 (
        echo.
        echo ERROR: Failed to install dependencies
        echo Please run manually: pip install -r requirements.txt
        pause
        exit /b 1
    )
)

echo.
echo Starting application...
echo.

REM Launch the GUI application
python heatmap_updater_gui.py

REM Keep window open if there was an error
if errorlevel 1 (
    echo.
    echo ERROR: Application exited with an error
    pause
)
