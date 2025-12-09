# AVL Drive Heatmap Tool

This repository contains the AVL Drive Heatmap Tool for evaluating drive operation modes.

## Recent Fix

### Operation Mode Status Evaluation Issue

**Issue:** The evaluation sheet's Operation Mode Summary was not correctly evaluating sub-operation statuses for **ALL operation modes**. When sub-operations had RED status, their parent operation modes were not showing RED status. This affected all 13 parent modes and their 42+ sub-operations.

**Solution:** Fixed the `InferParentMode` function in the VBA code to match operation codes based on the first 4 digits instead of all 8 digits. This ensures proper status aggregation across all operation mode families (Drive Away, Acceleration, Gear shift, TCC control, etc.).

## Files

- **AVLDrive_Heatmap_Tool version3.1.xlsm** - Main Excel workbook with VBA macros
- **VBA_CODE_FIX.md** - Detailed documentation of the fix including technical details
- **apply_vba_fix.py** - Python utility script to check if the fix has been applied

## How to Apply the Fix

### Prerequisites

- Microsoft Excel or LibreOffice Calc with macro support enabled
- OR Python 3 with `oletools` package for verification

### Steps

1. **Read the fix documentation:**
   ```bash
   cat VBA_CODE_FIX.md
   ```

2. **Check if fix is needed:**
   ```bash
   python3 apply_vba_fix.py
   ```

3. **Apply the fix manually:**
   - Open `AVLDrive_Heatmap_Tool version3.1.xlsm`
   - Press `Alt+F11` to open VBA Editor
   - Find the `Evaluation` module
   - Locate the `InferParentMode` function
   - Replace it with the fixed version from `VBA_CODE_FIX.md`
   - Save and close

4. **Verify the fix:**
   - Run the evaluation macro
   - Check that sub-operations are correctly grouped under their parent operation modes:
     - 10101300, 10101100, 10102400 → 10100000 (Drive Away)
     - 10120100, 10120200, 10120300 → 10120000 (Acceleration)
     - 10092300, 10093200, 10098200 → 10090000 (Gear shift)
   - Verify that if sub-operations have RED status, their parents also show RED

## Technical Details

The tool evaluates automotive drive operation modes based on test results. Operation modes use an 8-digit code system where the first 4 digits identify the mode family:

- **1010**xxxx - Drive away operations (Creep, Standing start, Rolling start)
- **1012**xxxx - Acceleration operations (Full load, Constant load, Load increase/decrease)
- **1009**xxxx - Gear shift operations (Upshift, Downshift, Maneuvering)
- **1046**xxxx - TCC control operations (Lock up, Controlled slip, Release)
- **1003**xxxx, **1004**xxxx, **1007**xxxx, **1008**xxxx, **1001**xxxx, **1002**xxxx, **1014**xxxx, **1043**xxxx, **1045**xxxx - Other mode families

The fix ensures that **all 42+ sub-operations** are correctly associated with their parent operations (13 total) for accurate status aggregation in the Operation Mode Summary.

## Requirements

To use the Python verification script:

```bash
pip install oletools
```

## Support

For issues or questions about the fix, refer to the detailed documentation in `VBA_CODE_FIX.md`.
