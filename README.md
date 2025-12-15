# AVLDrive HeatMap Tool - Status Update Script

This repository contains a Python script to automatically update the HeatMap Sheet status column based on evaluation results.

## Overview

The script reads evaluation results from the "Evaluation Results" sheet and updates the Status column (Column R) in the "HeatMap Sheet" according to these rules:

- **Sub-operations**: Filled with colored dots (●)
  - 🔴 Red dot for RED status (failure)
  - 🟡 Yellow dot for YELLOW status (warning/marginal)
  - 🟢 Green dot for GREEN status (pass)

- **Parent operations**: Filled with status text
  - "NOK" (red) for RED status
  - "acceptable" (yellow) for YELLOW status  
  - "OK" (green) for GREEN status
  - Parent status is automatically calculated as the worst status of all its sub-operations

## Prerequisites

- Python 3.6 or higher
- openpyxl library

## Installation

1. Install the required Python library:
   ```bash
   pip install openpyxl
   ```

## Usage

1. Ensure the Excel file `AVLDrive_Heatmap_Tool version3.2.xlsm` is in the same directory as the script

2. Run the script:
   ```bash
   python3 update_heatmap_status.py
   ```

3. The script will:
   - Read evaluation results from the "Evaluation Results" sheet
   - Update the Status column in "HeatMap Sheet"
   - Save the updated file (overwrites the original)

## How It Works

### Hierarchy Detection

The script automatically detects the operation hierarchy:
- **Parent operations**: OpCodes ending with 4 or more zeros (e.g., 10100000, 10030000)
- **Sub-operations**: More specific OpCodes (e.g., 10101300, 10030100)

### Status Calculation

1. **Sub-operations**: 
   - Matches OpCode from HeatMap Sheet with Evaluation Results sheet
   - Uses the Final Status from Column L of Evaluation Results
   - Fills Status column with colored dot (●) matching the evaluation status

2. **Parent operations**:
   - Aggregates statuses from all child sub-operations
   - Determines worst status (RED > YELLOW > GREEN)
   - Fills Status column with appropriate text and color

### Example

Given evaluations:
- OpCode 10030100 (At deceleration): YELLOW
- OpCode 10030200 (At constant speed): RED

The script will:
- Set 10030100 status to yellow dot (●)
- Set 10030200 status to red dot (●)
- Set parent 10030000 (Tip in) status to "NOK" (red) because worst sub-op is RED

## File Structure

- `update_heatmap_status.py` - Main Python script
- `AVLDrive_Heatmap_Tool version3.2.xlsm` - Excel file with HeatMap and Evaluation Results
- `README.md` - This documentation file

## Notes

- The script preserves VBA macros in the Excel file
- Original file is overwritten - consider backing up before running
- OpCodes not found in the HeatMap Sheet are skipped (expected behavior)
- N/A statuses from Evaluation Results are ignored

## Troubleshooting

**Error: File not found**
- Ensure the Excel file is in the same directory as the script
- Check the filename matches exactly: `AVLDrive_Heatmap_Tool version3.2.xlsm`

**Error: Sheet not found**
- Verify the Excel file contains "Evaluation Results" and "HeatMap Sheet" sheets
- Check sheet names match exactly (case-sensitive)

**No updates visible**
- Ensure Evaluation Results sheet has data in Column L (Final Status)
- Check that OpCodes in both sheets match
- Verify statuses are not N/A
