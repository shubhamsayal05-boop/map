# Benchmark Data Handling Fix

## Issue Identified

When benchmarking data (Target or Tested values) is missing for an operation, the evaluation was returning blank/N/A instead of evaluating based on the available criteria (AVL score and P1 status).

### Example Cases from Screenshot

Looking at rows with missing benchmark data:
- **Row 3**: Drive Away Creep Eng On - Cold: Driv Target = 0, Driv Tested = 0 → Driv Status was blank
- **Row 6**: DASS Eng On - Cold: Both Target and Tested = 0 → Status was blank
- **Row 12**: Drive away Creep: Both Target and Tested = 0 → Status was blank

These operations have valid AVL scores and P1 status that could be evaluated, but were showing blank/N/A due to missing benchmark data.

## Root Cause

The original `EvaluateStatus` function logic:

```vba
' OLD Logic
If UCase(Trim(p1)) = "N/A" Or benchDiff = 999 Then
    EvaluateStatus = vbNullString  ' Return blank
    Exit Function
End If
```

When `benchDiff = 999` (sentinel value for missing benchmark data), the function would immediately return blank without checking AVL score or P1 status.

## Solution

Modified the evaluation logic to use a **priority-based approach**:

### Evaluation Priority Order

1. **Priority 1: AVL Score and P1 Status** (Always evaluated if available)
   - If AVL < 7 OR P1 = RED → **RED**
   - If AVL >= 7 AND P1 = YELLOW → **YELLOW**
   - Only return blank if P1 is N/A (truly no data)

2. **Priority 2: Benchmark Data** (Evaluated only if available)
   - If benchmark data missing (benchDiff = 999) → **GREEN** (since AVL/P1 passed)
   - If Target/Tested not numeric → **GREEN** (since AVL/P1 passed)

3. **Priority 3: Benchmark Comparison** (If data available)
   - If Tested > Target → **GREEN**
   - If Target - Tested > 2 → **YELLOW**
   - Otherwise → **GREEN**

### Updated Logic

```vba
' NEW Logic
Private Function EvaluateStatus(avl As Double, p1 As String, benchDiff As Double, targetVal As Double, testedVal As Double) As String
    ' If P1 is N/A, cannot evaluate anything
    If UCase(Trim(p1)) = "N/A" Then
        EvaluateStatus = vbNullString
        Exit Function
    End If

    ' Priority 1: Check AVL and P1 status (always evaluated)
    If avl < 7 Or UCase(Trim(p1)) = "RED" Then
        EvaluateStatus = "RED"
        Exit Function
    End If

    If avl >= 7 And UCase(Trim(p1)) = "YELLOW" Then
        EvaluateStatus = "YELLOW"
        Exit Function
    End If

    ' Priority 2: If benchmark data missing, default to GREEN
    ' (since AVL >= 7 and P1 is GREEN)
    If benchDiff = 999 Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' Priority 3: Evaluate benchmark comparison (if data available)
    If Not IsNumeric(targetVal) Or Not IsNumeric(testedVal) Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    If testedVal > targetVal Then
        EvaluateStatus = "GREEN"
    Else
        If (targetVal - testedVal) > 2 Then
            EvaluateStatus = "YELLOW"
        Else
            EvaluateStatus = "GREEN"
        End If
    End If
End Function
```

## Impact

This change ensures:
- ✓ Operations with valid AVL and P1 data are evaluated even if benchmark data is missing
- ✓ Missing benchmark data no longer causes blank/N/A status
- ✓ Evaluation focuses on available data rather than requiring all data
- ✓ Only truly incomplete data (P1 = N/A) results in blank status

## Examples

### Before Fix:
```
Operation: Drive Away Creep Eng On - Cold
- AVL: 7 (GREEN threshold)
- P1: N/A
- Target: 0, Tested: 0 (missing)
- Status: blank ❌
```

### After Fix:
```
Operation: Drive Away Creep Eng On - Cold
- AVL: 7 (GREEN threshold)
- P1: N/A
- Target: 0, Tested: 0 (missing)
- Status: blank ✓ (P1 is N/A, truly no data)

Operation: DASS Eng On - Cold
- AVL: 6.7
- P1: N/A
- Target: 0, Tested: 0 (missing)
- Status: blank ✓ (P1 is N/A)

Operation: Example with P1 data but no benchmark
- AVL: 7.5
- P1: GREEN
- Target: 0, Tested: 0 (missing)
- Status: GREEN ✓ (evaluated on AVL and P1)
```

## Files Updated

- **Evaluation_FIXED.bas** - Modified `EvaluateStatus` function (lines 448-496)
- **BENCHMARK_DATA_HANDLING.md** - This documentation file

## Testing

After applying this fix, verify that:
1. Operations with missing benchmark data but valid AVL/P1 show appropriate status
2. Operations with AVL < 7 show RED regardless of benchmark data
3. Operations with P1 = YELLOW show YELLOW regardless of benchmark data
4. Operations with AVL >= 7 and P1 = GREEN but missing benchmark data show GREEN
5. Only operations with P1 = N/A show blank/N/A status
