# Car Selection Feature for Dynamic Evaluation

## Overview

This feature adds dynamic car selection capability to the evaluation process, allowing users to select Target and Tested cars from dropdown lists before running the evaluation.

## What's New

### Dropdown Controls
- **Location**: Columns W and X in Sheet1 (after column V as requested)
- **Target Car Dropdown**: Column W, Row 1 (light yellow background)
- **Tested Car Dropdown**: Column X, Row 1 (light green background)

### How It Works

1. **Initialize Dropdowns**: Run `InitializeCarSelectionDropdowns()` to populate the dropdowns with available car names from your data
2. **Select Cars**: Choose Target car from column W and Tested car from column X
3. **Run Evaluation**: Execute `EvaluateAVLStatus()` which now uses your selected cars
4. **View Results**: Results sheet shows which cars were used in the evaluation

## Files Added

### 1. CarSelection_Module.bas
Complete module for managing car selection functionality:
- `InitializeCarSelectionDropdowns()` - Sets up the dropdowns
- `ValidateCarSelections()` - Ensures valid selections before evaluation
- `GetTargetCarName()` - Returns selected target car name
- `GetTestedCarName()` - Returns selected tested car name
- `FindCarColumn()` - Locates column for a specific car
- `GetSelectedCarColumns()` - Returns column indices for both cars

### 2. Evaluation_WITH_CAR_SELECTION.bas
Modified evaluation module that uses dynamic car selection:
- All evaluation logic updated to read from selected car columns
- Results headers show which cars were evaluated
- Maintains all existing fixes (sub-operation matching, status evaluation, tolerance)

## Installation Steps

### Step 1: Import CarSelection Module

1. Open your Excel file
2. Press `Alt + F11` to open VBA Editor
3. Go to `File` → `Import File`
4. Select `CarSelection_Module.bas`
5. Module will appear in your Modules folder

### Step 2: Replace Evaluation Module

1. In VBA Editor, find existing "Evaluation" module
2. Right-click → `Remove Evaluation`
3. Choose "No" when asked to export (unless you want to keep a backup)
4. Go to `File` → `Import File`
5. Select `Evaluation_WITH_CAR_SELECTION.bas`

### Step 3: Initialize Dropdowns

1. In Excel, press `Alt + F8` to open Macro dialog
2. Select `InitializeCarSelectionDropdowns`
3. Click `Run`
4. Dropdowns will appear in columns W and X with all available car names

## Usage

### Daily Workflow

1. **Open your data file**
2. **Select cars from dropdowns**:
   - Click on cell W1 and choose your Target car
   - Click on cell X1 and choose your Tested car
3. **Run evaluation**:
   - Press `Alt + F8`
   - Select `EvaluateAVLStatus`
   - Click `Run`
4. **Review results**:
   - Check "Evaluation Results" sheet
   - Column headers will show which cars were used
   - Example: "Driv Target (MY26_LB_1)" and "Driv Tested (22MY_5.7L)"

### Data Structure Requirements

The code assumes:
- **Car data starts in column H (column 8)**
- **Car names are in row 1**
- **Each car's data is in consecutive columns**

If your structure is different, you can adjust constants in `CarSelection_Module.bas`:
```vba
Const CAR_DATA_START_COL As Long = 8  ' Change if car data starts elsewhere
Const CAR_NAME_ROW As Long = 1        ' Change if car names are in different row
```

## Features

### Validation
- Checks that both Target and Tested cars are selected before evaluation
- Warns if same car is selected for both (but allows user to continue)
- Shows error if selected car cannot be found in data

### Flexibility
- Works with any number of cars in your dataset
- Automatically detects available car names from data sheet
- Dropdown lists update automatically when you re-initialize

### Results Clarity
- Evaluation results clearly show which cars were used
- Column headers include car names: "Driv Target (CarName)"
- Summary message displays both selected cars

## Troubleshooting

### "Could not find car" Error
**Problem**: Selected car name doesn't match any column header
**Solution**: 
- Check that car names in row 1 match dropdown selection exactly
- Re-run `InitializeCarSelectionDropdowns` to refresh dropdown lists
- Verify car data starts in expected column (default: column H)

### Dropdowns Not Appearing
**Problem**: Dropdowns not visible after initialization
**Solution**:
- Scroll to columns W and X to see them
- Check that Sheet1 exists and is the correct sheet
- Verify no protection is applied to the worksheet

### Wrong Data Being Used
**Problem**: Evaluation uses wrong columns
**Solution**:
- Verify car names in row 1 are unique and correct
- Check that `CAR_DATA_START_COL` constant matches your data layout
- Re-select cars from dropdowns to ensure selections are saved

### Evaluation Fails
**Problem**: Evaluation button doesn't work with new module
**Solution**:
- Ensure both modules (CarSelection and Evaluation) are imported
- Check there are no naming conflicts with existing modules
- Verify Excel has enabled macros (check Trust Center settings)

## Technical Notes

### Column Mapping
The modified evaluation reads data dynamically based on dropdown selections:

**Original (Fixed Columns)**:
```vba
drivTarget = wsSheet1.Cells(i, 8).Value   ' Always column H
drivTested = wsSheet1.Cells(i, 10).Value  ' Always column J
```

**New (Dynamic Columns)**:
```vba
drivTarget = wsSheet1.Cells(i, targetCol).Value  ' Based on selected Target car
drivTested = wsSheet1.Cells(i, testedCol).Value  ' Based on selected Tested car
```

### Dropdown Maintenance
Dropdowns are created with Excel's Data Validation feature:
- Type: List
- Source: Comma-separated car names from row 1
- Automatically populated by scanning columns H onwards

To refresh available cars:
- Run `InitializeCarSelectionDropdowns` again
- This re-scans the data and updates both dropdowns

## All Existing Fixes Included

This feature includes all previous fixes:

✓ **Fix 1**: Sub-operation matching (first 4 digits)
✓ **Fix 2**: Status evaluation logic (handles YELLOW with ≤35% correctly)
✓ **Fix 3**: Benchmark data handling with tolerance
  - AVL < 7 or P1 = RED → RED
  - AVL >= 7 and P1 = YELLOW → YELLOW
  - AVL >= 7 and P1 = GREEN + meeting benchmark → GREEN
  - AVL >= 7 and P1 = GREEN + NOT meeting → YELLOW
  - Tolerance: tested < target AND within 2 units → GREEN

## Example Scenarios

### Scenario 1: Compare Two Different Cars
```
Target Car: MY26_LB_1
Tested Car: 22MY_5.7L
Result: Evaluation compares 22MY_5.7L data against MY26_LB_1 benchmarks
```

### Scenario 2: Self-Comparison
```
Target Car: Current_Model
Tested Car: Current_Model
Result: Evaluation compares car against itself (useful for baseline checks)
```

### Scenario 3: Multiple Evaluations
```
First Run:  Target=CarA, Tested=CarB
Second Run: Target=CarA, Tested=CarC
Result: Can easily compare multiple cars against same target
```

## Support

For issues or questions:
1. Check that all files are imported correctly
2. Verify data structure matches expected format
3. Test with sample data first
4. Review VBA code comments for additional details

## Quick Reference

| Action | Macro Name | Location |
|--------|-----------|----------|
| Setup dropdowns | InitializeCarSelectionDropdowns | CarSelection module |
| Run evaluation | EvaluateAVLStatus | Evaluation module |
| Get target car | GetTargetCarName | CarSelection module |
| Get tested car | GetTestedCarName | CarSelection module |
| Validate selections | ValidateCarSelections | CarSelection module |

## Version History

- **v1.0**: Initial release with car selection feature
- Builds on previous fixes (sub-operation matching, status evaluation, tolerance)
