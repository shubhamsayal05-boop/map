# Fix for Drive Away Operation Mode Status Evaluation

## Problem Description

The evaluation sheet's Operation Mode Summary is not correctly evaluating sub-operation statuses. When sub-operations of Drive Away (codes starting with "1010") have RED status, the parent Drive Away operation mode should also show RED status, but currently it does not.

### Example Issue:
- Parent operation: `10100000` (Drive away)
- Sub-operations:
  - `10101300` Drive Away Creep Eng On - Status: **RED**
  - `10101100` DASS Eng On - Status: **RED**
  - `10102400` DA Rolling Start - Status: **RED**
- Expected: Drive Away operation mode summary should be **RED**
- Actual: Drive Away operation mode summary is not correctly evaluated

## Root Cause

The `InferParentMode` function in the `Evaluation` module tries to match operation codes by comparing the entire 8-character code as a prefix. This fails for sub-operations.

### Current Logic (Incorrect):
```vba
For Each k In modes.Keys
    If Len(code) >= Len(k) Then
        If Left$(code, Len(k)) = k Then  ' Compares all 8 digits
            InferParentMode = k
            Exit Function
        End If
    End If
Next k
```

**Problem:** `Left$("10101300", 8)` = `"10101300"` ≠ `"10100000"` ❌

## Solution

Match based on the **first 4 digits** since all operation modes follow the pattern `"10XX0000"` where `XX` identifies the mode:
- `1010` = Drive away
- `1012` = Acceleration  
- `1003` = Tip in
- etc.

### Fixed Logic:
```vba
' Match based on first 4 digits since all operation modes follow pattern "10XX0000"
' where XX identifies the mode (e.g., 1010 = Drive away)
For Each k In modes.Keys
    If Len(code) >= 4 And Len(k) >= 4 Then
        If Left$(code, 4) = Left$(k, 4) Then  ' Compares first 4 digits only
            InferParentMode = k
            Exit Function
        End If
    End If
Next k
```

**Now:** `Left$("10101300", 4)` = `"1010"` = `Left$("10100000", 4)` = `"1010"` ✓

## How to Apply the Fix

### Option 1: Manual Update (Recommended)

1. **Open the Excel file:**
   - Open `AVLDrive_Heatmap_Tool version3.1.xlsm` in Microsoft Excel or LibreOffice Calc

2. **Open VBA Editor:**
   - Press `Alt+F11` to open the VBA Editor

3. **Locate the Evaluation module:**
   - In the Project Explorer (left panel), find and double-click `Evaluation` under `Modules`

4. **Find the InferParentMode function:**
   - Press `Ctrl+F` to open Find dialog
   - Search for `InferParentMode`
   - Scroll to the `Private Function InferParentMode` (around line 303)

5. **Replace the function:**
   - Select and delete the entire function from `Private Function InferParentMode` to its `End Function`
   - Paste the complete fixed function provided below

6. **Save the file:**
   - Press `Ctrl+S` to save
   - Close the VBA Editor

### Complete Fixed Function

Replace the entire `InferParentMode` function with this code:

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

## Verification

After applying the fix:

1. Run the evaluation macro
2. Check the Operation Mode Summary section
3. Verify that:
   - Sub-operations like `10101300`, `10101100`, `10102400` are evaluated
   - If all have RED status, Drive Away (`10100000`) should show RED
   - If any has RED status, Drive Away should show RED
   - If all are GREEN, Drive Away should show GREEN

## Changes Summary

**File:** `AVLDrive_Heatmap_Tool version3.1.xlsm`  
**Module:** `Evaluation.bas`  
**Function:** `InferParentMode`  
**Lines changed:** 311-315 (approximately)

**Before:**
```vba
If Len(code) >= Len(k) Then
    If Left$(code, Len(k)) = k Then
```

**After:**
```vba
If Len(code) >= 4 And Len(k) >= 4 Then
    If Left$(code, 4) = Left$(k, 4) Then
```

## Technical Details

### Operation Mode Hierarchy

All operation modes in the system use 8-digit codes with the pattern `10XX0000`:

| Code | First 4 Digits | Operation Mode |
|------|---------------|----------------|
| 10100000 | 1010 | Drive away |
| 10120000 | 1012 | Acceleration |
| 10030000 | 1003 | Tip in |
| 10040000 | 1004 | Tip out |
| 10070000 | 1007 | Deceleration |
| 10090000 | 1009 | Gear shift |
| 10080000 | 1008 | Constant speed |
| 10010000 | 1001 | Idle |
| 10020000 | 1002 | Engine start |
| 10140000 | 1014 | Engine shut off |
| 10460000 | 1046 | TCC control |
| 10430000 | 1043 | Cylinder deactivation |
| 10450000 | 1045 | Vehicle stationary |

Sub-operations share the first 4 digits with their parent operation. For example:
- `10101300` → First 4 digits: `1010` → Parent: Drive away (`10100000`)
- `10101100` → First 4 digits: `1010` → Parent: Drive away (`10100000`)
- `10102400` → First 4 digits: `1010` → Parent: Drive away (`10100000`)

### Status Evaluation Logic

The `BuildOperationModeSummary` function:
1. Iterates through all evaluated operation codes
2. For each code, checks if it's a primary operation mode or a sub-operation
3. If it's a sub-operation, uses `InferParentMode` to find the parent
4. Appends the status to the parent's status array
5. Evaluates the final status for each parent:
   - **RED** if any sub-operation is RED
   - **YELLOW** if >35% are YELLOW
   - **GREEN** if all are GREEN
   - **N/A** otherwise

With the fix, sub-operations will be correctly associated with their parents, ensuring accurate status aggregation.

---

**Version:** 1.0  
**Date:** 2025-11-21  
**Author:** GitHub Copilot Coding Agent
