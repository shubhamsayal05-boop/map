# AVL Drive Heatmap Tool

This repository contains the AVL Drive Heatmap Tool for evaluating drive operation modes.

## Recent Fix

### Operation Mode Status Evaluation Issue

**Issue:** The evaluation sheet's Operation Mode Summary was not correctly evaluating sub-operation statuses. When sub-operations of Drive Away (codes starting with "1010") had RED status, the parent Drive Away operation mode was not showing RED status.

**Solution:** Fixed the `InferParentMode` function in the VBA code to match operation codes based on the first 4 digits instead of all 8 digits.

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
   - Check that sub-operations (e.g., 10101300, 10101100, 10102400) are correctly grouped under their parent operation mode (10100000 - Drive Away)
   - Verify that if sub-operations have RED status, the parent also shows RED

## Technical Details

The tool evaluates automotive drive operation modes based on test results. Operation modes use an 8-digit code system where the first 4 digits identify the mode family:

- **1010**xxxx - Drive away operations
- **1012**xxxx - Acceleration operations  
- **1003**xxxx - Tip in operations
- etc.

The fix ensures that sub-operations are correctly associated with their parent operations for accurate status aggregation in the Operation Mode Summary.

## Requirements

To use the Python verification script:

```bash
pip install oletools
```

## Support

For issues or questions about the fix, refer to the detailed documentation in `VBA_CODE_FIX.md`.
