# Quick Start Guide

## Problem

The Operation Mode Summary in the evaluation sheet was not correctly evaluating sub-operation statuses for **ALL operation modes**. When sub-operations had RED status, their parent operation modes were not showing RED. This affects all 13 parent modes (Drive Away, Acceleration, Gear shift, etc.) and their 42+ sub-operations.

## Solution Provided

This repository now contains:

1. **VBA_CODE_FIX.md** - Detailed technical documentation
2. **apply_vba_fix.py** - Verification script
3. **README.md** - General documentation

## How to Fix (3 Simple Steps)

### Step 1: Verify the Issue

```bash
python3 apply_vba_fix.py
```

This will show you if the fix needs to be applied.

### Step 2: Apply the Fix

1. Open `AVLDrive_Heatmap_Tool version3.1.xlsm` in Excel
2. Press `Alt+F11` (opens VBA Editor)
3. Double-click `Evaluation` module in the left panel
4. Press `Ctrl+F` and search for `InferParentMode`
5. Replace the function with this code:

```vba
Private Function InferParentMode(code As String, modes As Object) As String
    If modes.Exists(code) Then
        InferParentMode = code
        Exit Function
    End If

    Dim k As Variant
    ' Match based on first 4 digits since all operation modes follow pattern "10XX0000"
    ' where XX identifies the mode (e.g., 1010 = Drive away)
    For Each k In modes.Keys
        If Len(code) >= 4 And Len(k) >= 4 Then
            If Left$(code, 4) = Left$(k, 4) Then
                InferParentMode = k
                Exit Function
            End If
        End If
    Next k

    InferParentMode = ""
End Function
```

6. Save (`Ctrl+S`) and close VBA Editor

### Step 3: Verify the Fix

Run the verification script again:

```bash
python3 apply_vba_fix.py
```

It should now show: "✓ The fix has ALREADY been applied!"

## What Changed?

**Before:**
```vba
If Left$(code, Len(k)) = k Then  ' Compares all 8 digits
```

**After:**
```vba
If Left$(code, 4) = Left$(k, 4) Then  ' Compares only first 4 digits
```

This ensures **ALL** sub-operations are correctly matched to their parent operations based on the first 4 digits. Examples:
- 10101300 (Creep) → 10100000 (Drive away) via "1010"
- 10120100 (Full load) → 10120000 (Acceleration) via "1012"
- 10092300 (Power-on upshift) → 10090000 (Gear shift) via "1009"

## Testing

After applying the fix:

1. Run your evaluation macro
2. Check the Operation Mode Summary
3. Verify that parent operations show correct status based on their sub-operations:
   - Drive Away (10100000) shows RED if its sub-operations are RED
   - Acceleration (10120000) shows RED if its sub-operations are RED
   - Gear shift (10090000) shows RED if its sub-operations are RED
   - Same logic applies to all 13 parent operation modes

## Need More Information?

- **Technical details:** See `VBA_CODE_FIX.md`
- **General info:** See `README.md`
- **Script help:** Run `python3 apply_vba_fix.py --help`

## Requirements

- Excel or LibreOffice with macro support
- Python 3 with `oletools` for verification (optional)

```bash
pip install oletools
```
