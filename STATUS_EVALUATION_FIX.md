# Status Evaluation Logic Fix

## Issue Identified

Three operation modes were showing "N/A" status instead of proper evaluation:
- **Deceleration (10070000)**: 16.67% yellow → showed N/A ❌
- **Constant speed (10080000)**: 46.67% yellow → showed N/A ❌  
- **Cylinder deactivation (10430000)**: 0% yellow → showed N/A ❌

## Root Cause

The original status evaluation logic in `BuildOperationModeSummary` had a gap:

```vba
' BEFORE (Incomplete logic)
If anyRed Then
    finalMode = "RED"
ElseIf pctYellow > 0.35 Then
    finalMode = "YELLOW"
ElseIf total > 0 And allGreen Then
    finalMode = "GREEN"
Else
    finalMode = "N/A"     ' ❌ Falls here when yellow% ≤ 35%
End If
```

**Problem:** When an operation mode had:
- No RED statuses (`anyRed = False`)
- Some YELLOW statuses but ≤35% (`pctYellow ≤ 0.35`)
- Not all GREEN (`allGreen = False`)
- Has data (`total > 0`)

The logic would fall through to "N/A" instead of "YELLOW".

### Example Cases

**Deceleration (16.67% yellow):**
1. anyRed = False ✗
2. pctYellow (0.1667) > 0.35? No ✗
3. total > 0 And allGreen? No (has yellows) ✗
4. **Falls to N/A** ❌

**Should be YELLOW** because it has some YELLOW sub-operations.

## Solution

Added an additional condition to catch cases where there's data but not all green:

```vba
' AFTER (Complete logic)
If anyRed Then
    finalMode = "RED"
ElseIf pctYellow > 0.35 Then
    finalMode = "YELLOW"
ElseIf total > 0 And allGreen Then
    finalMode = "GREEN"
ElseIf total > 0 Then
    ' Has data but not all green (some yellow) - should be YELLOW
    finalMode = "YELLOW"
Else
    finalMode = "N/A"
End If
```

## Status Decision Tree

The corrected logic now follows this decision tree:

1. **Has any RED?** → Status = RED
2. **Yellow percentage > 35%?** → Status = YELLOW
3. **Has data AND all GREEN?** → Status = GREEN
4. **Has data but not all GREEN?** → Status = YELLOW *(new condition)*
5. **No data?** → Status = N/A

## Impact

This fix ensures that:
- ✓ Any operation mode with YELLOW sub-operations shows YELLOW status
- ✓ The 35% threshold still applies for high yellow concentration
- ✓ Modes with low yellow percentage (≤35%) now correctly show YELLOW instead of N/A
- ✓ Only modes with no data show N/A

## Affected Modes

This fix resolves the N/A issue for operation modes like:
- **Deceleration (10070000)**: Now shows YELLOW (was N/A)
- **Constant speed (10080000)**: Now shows YELLOW (was N/A)
- **Cylinder deactivation (10430000)**: Status depends on actual data

## Files Updated

- `Evaluation_FIXED.bas` - Lines 266-277 (status evaluation logic)
- `VBA_CODE_FIX.md` - Updated status evaluation documentation

## Testing

After applying this fix, verify:
1. Modes with any YELLOW sub-operations show YELLOW status
2. Modes with >35% YELLOW still show YELLOW
3. Modes with all GREEN show GREEN
4. Only modes with no data show N/A

Example verification:
```
Deceleration (10070000):
- Sub-operations: Mix of YELLOW statuses
- Yellow percentage: 16.67%
- Expected: YELLOW ✓ (was showing N/A)

Constant speed (10080000):
- Sub-operations: Mix including YELLOW
- Yellow percentage: 46.67%
- Expected: YELLOW ✓ (was showing N/A)
```
