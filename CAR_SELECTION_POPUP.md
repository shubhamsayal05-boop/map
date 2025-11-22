# Car Selection Popup Dialog Feature

## Overview

This feature allows users to dynamically select Target and Tested cars through a popup dialog that appears when running the evaluation. This provides a cleaner, more intuitive interface than the previous dropdown approach.

## How It Works

1. User runs `EvaluateAVLStatus` macro
2. **Popup dialog appears** with prompts for Target and Tested car selection
3. User enters/selects car names from the list shown
4. Evaluation runs with the selected cars' data
5. Results sheet shows which cars were evaluated

## Installation

### Step 1: Import CarSelectionForm Module

1. Open your Excel workbook
2. Press `Alt+F11` to open VBA Editor
3. Go to `File` → `Import File`
4. Select `CarSelectionForm_Module.bas`
5. Click `Open`

### Step 2: Import or Replace Evaluation Module

1. If you have an existing Evaluation module:
   - Right-click on it in the VBA Project Explorer
   - Select `Remove Evaluation`
   - Click `No` when asked to export (unless you want a backup)

2. Import the new module:
   - Go to `File` → `Import File`
   - Select `Evaluation_WITH_POPUP.bas`
   - Click `Open`

### Step 3: Save and Test

1. Press `Ctrl+S` to save the workbook
2. Close VBA Editor (or keep open for debugging)
3. Press `Alt+F8` to open Macros dialog
4. Run `EvaluateAVLStatus`
5. Popup should appear asking for car selection

## Usage

### Running Evaluation

**Method 1: Using Macro Dialog**
1. Press `Alt+F8`
2. Select `EvaluateAVLStatus` from the list
3. Click `Run`
4. Popup appears with car selection prompts

**Method 2: Assign to Button (Recommended)**
1. Go to `Developer` tab → `Insert` → `Button` (Form Control)
2. Draw button on worksheet
3. Select `EvaluateAVLStatus` macro
4. Click `OK`
5. Now clicking button shows popup and runs evaluation

**Method 3: Keyboard Shortcut**
1. Press `Alt+F8`
2. Select `EvaluateAVLStatus`
3. Click `Options`
4. Assign shortcut key (e.g., `Ctrl+Shift+E`)
5. Click `OK`

### Selecting Cars

When the popup appears:

1. **Target Car Selection:**
   - Dialog shows list of available cars
   - Enter the exact name of the Target car
   - Click `OK`

2. **Tested Car Selection:**
   - Dialog shows list of available cars again
   - Enter the exact name of the Tested car
   - Click `OK`

3. **Evaluation Runs:**
   - If both selections valid → evaluation proceeds
   - Results appear in "Evaluation Results" sheet
   - Headers show selected car names

### Example Workflow

```
1. User clicks "Run Evaluation" button
2. Popup #1: "Available cars: MY26_LB_1, 22MY_5.7L, CarA, CarB"
   → User enters: MY26_LB_1
3. Popup #2: "Available cars: MY26_LB_1, 22MY_5.7L, CarA, CarB"
   → User enters: 22MY_5.7L
4. Evaluation runs comparing 22MY_5.7L (Tested) against MY26_LB_1 (Target)
5. Results show headers like "Driv Target (MY26_LB_1)" and "Driv Tested (22MY_5.7L)"
```

## Features

### Automatic Car Detection

- Scans row 1 of data sheet for car names
- Starts from column H (configurable)
- Shows all available cars in selection prompt
- Works with any number of cars

### Validation

- Checks if entered car names are valid
- Shows error message if car not found
- Warns if same car selected for both Target and Tested
- Allows user to cancel at any point

### User-Friendly

- Simple InputBox interface (built into Excel)
- Shows available cars in each prompt
- Clear error messages
- No complex form controls

## Data Requirements

### Data Sheet Structure

Your data sheet must have:

1. **Sheet Name:** "Sheet1" (configurable in code)
2. **Car Names:** Row 1, starting from column H onwards
3. **Car Data:** Each car's data in consecutive columns below its name

Example:
```
Row 1:  | Op Code | Operation | ... | CarA | CarB | CarC | ...
        | (Col A) | (Col B)   | ... | (H)  | (I)  | (J)  | ...
```

### Car Name Format

- Car names can contain any characters
- Case-sensitive matching
- Leading/trailing spaces are trimmed
- Must be unique

## Troubleshooting

### "No car names found" Error

**Cause:** No car names detected in row 1

**Solutions:**
- Ensure car names are in row 1
- Check they start from column H (or your configured start column)
- Verify cells are not empty
- Check sheet name matches "Sheet1" (or your configured name)

### "Invalid car name" Error

**Cause:** Entered car name doesn't match available cars exactly

**Solutions:**
- Copy-paste car name from the list shown in dialog
- Check for extra spaces or typos
- Ensure case matches exactly
- Verify car name is in row 1 of data sheet

### "Could not find data columns" Error

**Cause:** Selected car names found in popup but not in worksheet

**Solutions:**
- Close and reopen workbook
- Verify car names still exist in row 1
- Check worksheet hasn't been modified
- Re-run evaluation from beginning

### Popup Doesn't Appear

**Cause:** Macro not running or error in code

**Solutions:**
- Check macro security settings (File → Options → Trust Center → Macro Settings)
- Enable macros when opening workbook
- Open VBA Editor (Alt+F11) and check for compile errors
- Verify both modules imported correctly

### Wrong Data Evaluated

**Cause:** Column mapping issue or incorrect selection

**Solutions:**
- Verify you entered correct car names
- Check car names in results headers match your selection
- Ensure car data is in correct columns below names
- Review data sheet structure

## Customization

### Changing Data Sheet Name

Edit `CarSelectionForm_Module.bas`:

```vba
' Line ~18
Private Const DATA_SHEET_NAME As String = "YourSheetName"
```

### Changing Car Data Start Column

Edit `CarSelectionForm_Module.bas`:

```vba
' Line ~19
Private Const CAR_DATA_START_COL As Integer = 10  ' Column J instead of H
```

### Adding Dropdown List (Advanced)

For a true dropdown experience, you would need to create a UserForm:

1. In VBA Editor: Insert → UserForm
2. Add two ComboBox controls for Target and Tested
3. Add OK and Cancel buttons
4. Modify `ShowCarSelectionDialog()` to show the form
5. Populate ComboBoxes with car names

This requires more advanced VBA programming. The current InputBox approach is simpler and works well for most use cases.

## Technical Details

### Module Structure

**CarSelectionForm_Module.bas:**
- `ShowCarSelectionDialog()` - Main function to show prompts
- `GetSelectedTargetCar()` - Returns Target car name
- `GetSelectedTestedCar()` - Returns Tested car name
- `GetSelectedCarColumns()` - Returns column indices for both cars
- `GetAvailableCarNames()` - Scans worksheet for car names
- `IsCarNameValid()` - Validates entered car name
- `FindCarColumn()` - Locates data column for a car

**Evaluation_WITH_POPUP.bas:**
- Modified `EvaluateAVLStatus()` to call popup first
- Uses selected car columns instead of fixed columns
- All other evaluation logic unchanged

### Integration Points

1. **Evaluation Start:**
   ```vba
   ' Show car selection dialog
   If Not ShowCarSelectionDialog() Then
       MsgBox "Evaluation cancelled."
       Exit Sub
   End If
   ```

2. **Get Selections:**
   ```vba
   Dim cols As Variant
   cols = GetSelectedCarColumns()
   targetCol = cols(0)
   testedCol = cols(1)
   ```

3. **Read Data:**
   ```vba
   ' Instead of fixed column
   drivTarget = wsSheet1.Cells(i, targetCol).Value
   drivTested = wsSheet1.Cells(i, testedCol).Value
   ```

## Comparison: Popup vs Dropdown Approach

| Aspect | Popup Dialog | Dropdown (W/X) |
|--------|--------------|----------------|
| **User Interface** | Clean, no worksheet clutter | Dropdowns visible in sheet |
| **Workflow** | One-click, all in popup | Two steps: select, then run |
| **Error Prevention** | Must select before running | Can run without selection |
| **Flexibility** | Standard Excel pattern | Custom implementation |
| **Setup** | No setup needed | Must run initialization |
| **Visual Feedback** | Dialog always appears | Dropdowns must be found |
| **Code Complexity** | Simpler (InputBox) | More complex (Data Validation) |

## Advantages of Popup Approach

1. **Cleaner Interface** - No dropdown cells in worksheet
2. **Better User Experience** - Standard dialog pattern
3. **Prevents Errors** - Can't forget to select cars
4. **Simpler Code** - Uses built-in InputBox
5. **Easier Maintenance** - No worksheet cells to manage
6. **More Intuitive** - Users expect dialogs for selections

## Migration from Dropdown Approach

If you previously used the dropdown approach (columns W/X):

1. **Remove Old Modules:**
   - Remove `CarSelection_Module`
   - Remove `Evaluation_WITH_CAR_SELECTION`

2. **Clean Up Worksheet:**
   - Delete dropdowns in columns W and X (if present)
   - Remove any formatting in those columns

3. **Import New Modules:**
   - Import `CarSelectionForm_Module.bas`
   - Import `Evaluation_WITH_POPUP.bas`

4. **Test:**
   - Run `EvaluateAVLStatus`
   - Popup should appear instead of using dropdowns
   - Functionality should be identical

## Support

For issues or questions:

1. Check this documentation
2. Review troubleshooting section
3. Verify data sheet structure
4. Check VBA Editor for errors (Alt+F11)
5. Contact repository maintainer

## Version History

- **v2.0** (2025-11-22) - Popup dialog approach
- **v1.0** (2025-11-22) - Dropdown approach (columns W/X)
