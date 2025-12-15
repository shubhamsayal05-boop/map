#!/bin/bash
# Launch script for HeatMap Updater GUI (Linux/Mac)

echo "==========================================="
echo "AVLDrive HeatMap Updater - Starting..."
echo "==========================================="
echo ""

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed or not in PATH"
    echo "Please install Python 3.7 or higher from https://www.python.org"
    read -p "Press Enter to exit..."
    exit 1
fi

# Check Python version
PYTHON_VERSION=$(python3 --version 2>&1 | awk '{print $2}')
echo "Python version: $PYTHON_VERSION"

# Check if openpyxl is installed
if ! python3 -c "import openpyxl" &> /dev/null; then
    echo ""
    echo "WARNING: openpyxl is not installed"
    
    # Check if pip3 is available
    if ! command -v pip3 &> /dev/null; then
        echo "ERROR: pip3 is not installed"
        echo "Please install pip3 or install openpyxl manually"
        read -p "Press Enter to exit..."
        exit 1
    fi
    
    echo "Attempting to install dependencies..."
    echo ""
    pip3 install -r requirements.txt
    
    if [ $? -ne 0 ]; then
        echo ""
        echo "ERROR: Failed to install dependencies"
        echo "Please run manually: pip3 install -r requirements.txt"
        read -p "Press Enter to exit..."
        exit 1
    fi
fi

echo ""
echo "Starting application..."
echo ""

# Launch the GUI application
python3 heatmap_updater_gui.py

# Keep window open if there was an error
if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Application exited with an error"
    read -p "Press Enter to exit..."
fi
