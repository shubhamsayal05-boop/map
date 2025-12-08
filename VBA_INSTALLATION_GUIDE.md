# VBA Installation Guide - Update HeatMap Status Button

This guide explains how to add the VBA code to your Excel workbook to enable the "Update HeatMap Status" button functionality.

## Overview

The VBA code will automate the process of:
1. Reading evaluation results from the "Evaluation Results" sheet
2. Grouping them by Operation Code (Op Code)
3. Determining the worst status for each operation (RED > YELLOW > GREEN)
4. Updating the Status column (Column R) in the "HeatMap Sheet"

## Installation Steps

### Step 1: Open the VBA Editor

1. Open `AVLDrive_Heatmap_Tool version3.2.xlsm` in Excel
2. Press `Alt + F11` to open the VBA Editor
3. If prompted about macros, enable them

### Step 2: Add a New Module

1. In the VBA Editor, go to **Insert** → **Module**
2. A new module (e.g., "Module1") will appear in the Project Explorer

### Step 3: Copy the VBA Code

1. Open the file `UpdateHeatMapStatus.vba` in a text editor
2. Copy **ALL** the code from the file
3. Paste it into the new module in the VBA Editor

### Step 4: Link the Button to the Macro

1. Close the VBA Editor (or press `Alt + Q`)
2. Return to the "HeatMap Sheet" in Excel
3. Right-click on the **"Update HeatMap Status"** button
4. Select **Assign Macro...**
5. In the dialog, select **UpdateHeatMapStatus** from the list
6. Click **OK**

### Step 5: Test the Button

1. Make sure you're on the "HeatMap Sheet"
2. Click the **"Update HeatMap Status"** button
3. The macro will:
   - Read all evaluation data
   - Update the Status column (Column R)
   - Show a summary message with the number of updates

## Expected Results

After clicking the button, you should see:
- **Status column (Column R)** populated with RED, YELLOW, or GREEN values
- A message box showing:
  - Number of operations updated
  - Number of operations with no matching evaluation

Example:
```
Update Complete!

Updated: 28 operations
No match: 18 operations
```

## How It Works

### Status Priority Logic

The macro determines the worst status among multiple sub-operations:

1. **RED** = Priority 0 (Worst)
2. **YELLOW** = Priority 1 (Medium)
3. **GREEN** = Priority 2 (Good)
4. **N/A** or empty = Priority 3 (Ignored)

### Example

For Op Code `10101300` (Creep):
- Sub-operations: "RED", "N/A", "YELLOW", "N/A"
- **Final Status**: **RED** (worst status wins)

For Op Code `10040300` (Tip out at constant speed):
- Sub-operations: "GREEN"
- **Final Status**: **GREEN**

## Troubleshooting

### "Compile Error" when running the macro

**Solution**: Make sure you copied ALL the code, including the `Option Explicit` at the top.

### "Method 'Worksheets' failed"

**Solution**: Check that your sheets are named exactly:
- "Evaluation Results"
- "HeatMap Sheet"

### Status column not updating

**Solution**: 
1. Check that Column R is the Status column in HeatMap Sheet
2. Verify that evaluation data exists in "Evaluation Results" sheet
3. Check that Op Codes match between sheets

### Button not responding

**Solution**:
1. Right-click the button → Assign Macro
2. Select `UpdateHeatMapStatus` from the list
3. Click OK

## Additional Features

### Clear Status Column (Optional)

The VBA code also includes a utility function to clear all status values:

**To use it:**
1. Open VBA Editor (`Alt + F11`)
2. Press `F5` or go to **Run** → **Run Sub/UserForm**
3. Select `ClearHeatMapStatus`
4. Click **Run**

This will clear all values in the Status column.

## Code Structure

The VBA code consists of three main functions:

1. **`UpdateHeatMapStatus()`** - Main subroutine called by the button
2. **`DetermineWorstStatus()`** - Determines the worst status from multiple values
3. **`GetStatusPriority()`** - Assigns priority to each status color

## Maintenance

When the Evaluation Results are updated:
1. Simply click the **"Update HeatMap Status"** button again
2. The macro will re-read all data and update the Status column

## Performance Notes

- The macro disables screen updating during execution for better performance
- Typical execution time: < 1 second for the current data size
- The macro handles up to thousands of rows efficiently

## Support

If you encounter any issues:
1. Check that all sheets exist with correct names
2. Verify that columns are in the correct positions
3. Make sure macros are enabled in Excel

For the Python alternative, refer to `update_heatmap_status.py` and `README.md`.
