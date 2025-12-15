#!/usr/bin/env python3
"""
Launcher script for standalone HeatMap application
No Excel dependency - uses built-in SQLite database
"""

import sys
import os

# Check Python version
if sys.version_info < (3, 7):
    print("ERROR: Python 3.7 or higher is required")
    sys.exit(1)

# Check if in correct directory
script_dir = os.path.dirname(os.path.abspath(__file__))
os.chdir(script_dir)

print("=" * 60)
print("AVLDrive HeatMap Manager - Standalone Application")
print("=" * 60)
print()
print("Starting application...")
print()

try:
    # Import and run the standalone app
    from standalone_heatmap_app import main
    main()
except ImportError as e:
    print(f"ERROR: Missing module: {e}")
    print()
    print("This standalone app has no external dependencies!")
    print("All required modules are part of Python standard library.")
    sys.exit(1)
except Exception as e:
    print(f"ERROR: {e}")
    import traceback
    traceback.print_exc()
    sys.exit(1)
