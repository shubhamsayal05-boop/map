# HeatMap Status Update Feature

## Overview

This feature automatically transfers evaluation results from the "Evaluation Results" sheet to the "HeatMap Sheet", filling in the status dots (RED/YELLOW/GREEN) for all operations based on the evaluation.

## Features

- ✓ **One-click update** - Single button transfers all statuses
- ✓ **Color-coded dots** - Visual RED/YELLOW/GREEN indicators
- ✓ **Automatic matching** - Matches operation codes between sheets
- ✓ **Parent mode support** - Updates both sub-operations and parent operations
- ✓ **Bulk processing** - Updates all operations in one click
- ✓ **Error handling** - Clear messages if issues occur

## Installation

### Step 1: Import the Module

1. Open Excel workbook
2. Press `Alt + F11` to open VBA Editor
3. Go to **File → Import File**
4. Select `HeatMapUpdate_Module.bas`
5. Click **Open**

### Step 2: Create the Button

1. In VBA Editor, press `Alt + F8` (or go to Excel and press `Alt + F8`)
2. Select `CreateUpdateButton` from the macro list
3. Click **Run**
4. A button will be created on the HeatMap Sheet

## Usage

### Daily Workflow

1. **Run Evaluation**
   - Press `Alt + F8`
   - Run `EvaluateAVLStatus`
   - Select Target and Tested cars in popup dialogs
   - Wait for evaluation to complete

2. **Update HeatMap**
   - Go to **HeatMap Sheet**
   - Click **"Update HeatMap Status"** button
   - Wait for completion message

3. **Review Results**
   - Check that all operations now have colored dots in "Current Status P1" column
   - Verify colors match evaluation results

## What Gets Updated

### Sub-Operations
The "Final Status" from the evaluation results is transferred to the HeatMap Sheet:
- **Creep (10101300)** → Status dot in row for Creep
- **DASS Eng On (10101100)** → Status dot in row for DASS Eng On
- **DA Rolling Start (10102400)** → Status dot in row for DA Rolling Start
- And all other sub-operations...

### Parent Operations
The "Final Status" from the Operation Mode Summary is transferred:
- **Drive away (10100000)** → Status dot in row for Drive away
- **Acceleration (10120000)** → Status dot in row for Acceleration
- **Gear shift (10090000)** → Status dot in row for Gear shift
- And all other parent operations...

## Color Coding

| Status | Color | Symbol |
|--------|-------|--------|
| RED | Red (RGB 255, 0, 0) | ● |
| YELLOW | Orange/Yellow (RGB 255, 192, 0) | ● |
| GREEN | Green (RGB 0, 176, 80) | ● |
| N/A | Gray (RGB 166, 166, 166) | ● |

## Data Structure Requirements

### Evaluation Results Sheet

**Sub-Operations Section:**
- Column A: Op Code (e.g., "10101300")
- Column M (13): Final Status ("RED", "YELLOW", "GREEN")

**Operation Mode Summary Section:**
- Column F (6): Op Code (e.g., "10100000")
- Column I (9): Final Status ("RED", "YELLOW", "GREEN")

### HeatMap Sheet

- Column A: Op Code
- Column C: Current Status P1 (where dots will be filled)
- Column D: Current Status P2 (preserved, not modified)
- Column E: Current Status P3 (preserved, not modified)

## Example

### Before Update

```
HeatMap Sheet - Current Status P1 column is empty:

Op Code    Operation                    Current Status
                                        P1    P2    P3
10101300   Drive Away Creep Eng On     
10101100   DASS Eng On                 
10100000   Drive away                  
```

### After Update

```
HeatMap Sheet - Current Status P1 column filled with colored dots:

Op Code    Operation                    Current Status
                                        P1    P2    P3
10101300   Drive Away Creep Eng On     ●(RED)
10101100   DASS Eng On                 ●(RED)
10100000   Drive away                  ●(RED)
```

## Troubleshooting

### Button Not Working

**Problem:** Clicking button does nothing or shows error

**Solution:**
1. Verify `HeatMapUpdate_Module.bas` is imported
2. Check macro security settings (File → Options → Trust Center → Macro Settings → Enable all macros)
3. Re-create button by running `CreateUpdateButton` macro

### No Operations Updated

**Problem:** Message shows "Operations updated: 0"

**Solutions:**
1. **Verify evaluation ran successfully** - Check "Evaluation Results" sheet has data
2. **Check operation codes match** - Op codes in HeatMap must match Evaluation Results
3. **Verify column structure** - Ensure columns are in correct positions

### Wrong Columns Updated

**Problem:** Dots appear in wrong columns

**Solution:**
- The code updates Column C (Current Status P1)
- If your HeatMap has different structure, modify the `UpdateOperationStatus` function:
  ```vba
  Set statusCell = wsHeatMap.Cells(i, 3) ' Change 3 to your column number
  ```

### Some Operations Not Updated

**Problem:** Some operations missing from update

**Solutions:**
1. **Check operation code format** - Must be exact match (e.g., "10101300" not "101013")
2. **Verify operation exists in HeatMap** - Some operations may not be in HeatMap
3. **Check for hidden rows** - Unhide all rows in HeatMap Sheet

## Customization

### Change Button Position

Edit the `CreateUpdateButton` macro:

```vba
Set btn = wsHeatMap.Buttons.Add(Left, Top, Width, Height)
' Example: Move to cell E2
Set btn = wsHeatMap.Buttons.Add(300, 50, 150, 30)
```

### Change Button Text

Edit the `CreateUpdateButton` macro:

```vba
.Caption = "Your Custom Text Here"
```

### Change Status Column

Edit the `UpdateOperationStatus` function:

```vba
Set statusCell = wsHeatMap.Cells(i, ColumnNumber) ' Change column number
```

### Change Colors

Edit the `UpdateOperationStatus` function color section:

```vba
Case "RED"
    statusCell.Font.Color = RGB(R, G, B) ' Change RGB values
```

## Advanced Features

### Automatic Update After Evaluation

To automatically update HeatMap after evaluation completes, add this to the end of `EvaluateAVLStatus` macro:

```vba
' Auto-update HeatMap
Call UpdateHeatMapStatus
```

### Update Specific Section Only

Create a custom macro to update only Drivability or Responsiveness:

```vba
Sub UpdateDrivabilityOnly()
    ' Modify UpdateHeatMapStatus to filter by section
    ' Add section parameter to function
End Sub
```

## Technical Notes

### Performance

- Processes ~50-100 operations in 1-2 seconds
- Uses Application.ScreenUpdating = False for speed
- Shows progress in status bar

### Compatibility

- Works with Excel 2010 and later
- Compatible with both .xls and .xlsm formats
- No external dependencies

### Data Validation

- Checks if worksheets exist
- Validates operation codes are numeric
- Skips blank rows automatically
- Handles missing data gracefully

## Support

For issues or questions:
1. Check this documentation first
2. Verify all prerequisites are met
3. Review error messages carefully
4. Check that evaluation ran successfully before updating

## Quick Reference

| Action | Macro Name | Shortcut |
|--------|-----------|----------|
| Create button | CreateUpdateButton | Alt+F8 |
| Update HeatMap | UpdateHeatMapStatus | Click button |
| View code | - | Alt+F11 |

## Change Log

### Version 1.0 (2025-11-22)
- Initial release
- Basic status transfer functionality
- Button creation utility
- Support for sub-operations and parent operations
- Color-coded status dots
