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

Modified the evaluation logic to match the specification:

### Evaluation Rules

1. **AVL > 7 AND P1 = GREEN AND meeting benchmark** → **GREEN** (OK)
2. **AVL > 7 AND P1 = GREEN AND NOT meeting benchmark** → **YELLOW** (Acceptable, improve if possible)
3. **AVL > 7 AND P1 = YELLOW AND meeting benchmark** → **YELLOW** (Acceptable, improve if possible)
4. **AVL > 7 AND P1 = YELLOW AND NOT meeting benchmark** → **YELLOW** (Acceptable, improve if possible)
5. **AVL < 7 OR P1 = RED** → **RED** (NOK improve or buy off)
6. **AVL < 7 OR P1 = RED AND meeting benchmark** → **RED** (still NOK improve or buy off)
7. **If no benchmark data** → ignore benchmark and evaluate on AVL and P1 only

### Key Points

- "Meeting benchmark" means: `tested >= target` (meeting or exceeding the target)
- If AVL < 7 or P1 = RED → always RED (regardless of benchmark)
- If AVL >= 7 and P1 = YELLOW → always YELLOW (regardless of benchmark)
- If AVL >= 7 and P1 = GREEN:
  - With benchmark: meeting → GREEN, not meeting → YELLOW
  - Without benchmark: GREEN

### Updated Logic

```vba
' NEW Logic - Following specification
Private Function EvaluateStatus(avl As Double, p1 As String, benchDiff As Double, targetVal As Double, testedVal As Double) As String
    ' If P1 is N/A, cannot evaluate anything
    If UCase(Trim(p1)) = "N/A" Then
        EvaluateStatus = vbNullString
        Exit Function
    End If

    ' Rule 5 & 6: If AVL < 7 OR P1 = RED → Always RED (regardless of benchmark)
    If avl < 7 Or UCase(Trim(p1)) = "RED" Then
        EvaluateStatus = "RED"
        Exit Function
    End If

    ' Rule 3 & 4: If P1 = YELLOW → Always YELLOW (regardless of benchmark)
    If avl >= 7 And UCase(Trim(p1)) = "YELLOW" Then
        EvaluateStatus = "YELLOW"
        Exit Function
    End If

    ' At this point: AVL >= 7 AND P1 = GREEN
    
    ' If benchmark data is missing, ignore it and evaluate on AVL/P1 only
    If benchDiff = 999 Then
        ' AVL >= 7 AND P1 = GREEN AND no benchmark data → GREEN
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' If benchmark values not numeric, ignore benchmark
    If Not IsNumeric(targetVal) Or Not IsNumeric(testedVal) Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' Rule 1 & 2: Benchmark data is available
    ' Meeting benchmark: tested >= target → GREEN
    ' Not meeting benchmark: tested < target → YELLOW
    If testedVal >= targetVal Then
        EvaluateStatus = "GREEN"
    Else
        EvaluateStatus = "YELLOW"
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

### Example 1: AVL > 7, P1 = GREEN, Meeting Benchmark
```
- AVL: 7.5
- P1: GREEN
- Target: 100, Tested: 105 (tested >= target)
- Status: GREEN ✓ (Rule 1: Meeting benchmark)
```

### Example 2: AVL > 7, P1 = GREEN, NOT Meeting Benchmark
```
- AVL: 7.5
- P1: GREEN
- Target: 100, Tested: 95 (tested < target)
- Status: YELLOW ✓ (Rule 2: Not meeting benchmark, acceptable but improve)
```

### Example 3: AVL > 7, P1 = YELLOW, Meeting Benchmark
```
- AVL: 7.5
- P1: YELLOW
- Target: 100, Tested: 105 (tested >= target)
- Status: YELLOW ✓ (Rule 3: P1 is YELLOW, always YELLOW)
```

### Example 4: AVL > 7, P1 = YELLOW, NOT Meeting Benchmark
```
- AVL: 7.5
- P1: YELLOW
- Target: 100, Tested: 95 (tested < target)
- Status: YELLOW ✓ (Rule 4: P1 is YELLOW, always YELLOW)
```

### Example 5: AVL < 7, P1 = GREEN
```
- AVL: 6.5
- P1: GREEN
- Target: 100, Tested: 105
- Status: RED ✓ (Rule 5: AVL < 7, always RED)
```

### Example 6: P1 = RED, Meeting Benchmark
```
- AVL: 7.5
- P1: RED
- Target: 100, Tested: 105 (meeting benchmark)
- Status: RED ✓ (Rule 6: P1 is RED, always RED)
```

### Example 7: No Benchmark Data
```
- AVL: 7.5
- P1: GREEN
- Target: 0, Tested: 0 (missing benchmark data)
- Status: GREEN ✓ (Rule 7: Ignore benchmark, evaluate on AVL and P1 only)
```

### Example 8: Truly No Data
```
- AVL: 7.0
- P1: N/A
- Target: 0, Tested: 0
- Status: blank ✓ (P1 is N/A, cannot evaluate)
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
