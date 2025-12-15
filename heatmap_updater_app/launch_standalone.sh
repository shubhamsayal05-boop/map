#!/bin/bash
# Launcher for Standalone HeatMap Application (Linux/Mac)

echo "============================================================"
echo "AVLDrive HeatMap Manager - Standalone Application"
echo "============================================================"
echo ""

# Get the directory where the script is located
SCRIPT_DIR="$( cd "$( dirname "${BASH_SOURCE[0]}" )" && pwd )"
cd "$SCRIPT_DIR"

# Check if Python is available
if ! command -v python3 &> /dev/null; then
    echo "ERROR: Python 3 is not installed or not in PATH"
    echo "Please install Python 3.7+ from https://www.python.org"
    read -p "Press Enter to exit..."
    exit 1
fi

echo "Starting standalone application..."
echo "(No external dependencies required)"
echo ""

# Launch the standalone application
python3 launch_standalone.py

if [ $? -ne 0 ]; then
    echo ""
    echo "ERROR: Application exited with an error"
    read -p "Press Enter to exit..."
fi
