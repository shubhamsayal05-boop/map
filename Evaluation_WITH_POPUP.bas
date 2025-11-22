Attribute VB_Name = "Evaluation"
Option Explicit

' Main entry: builds "Evaluation Results" sheet and summaries
' Now uses popup dialog for car selection
Public Sub EvaluateAVLStatus()
    Dim wsSheet1 As Worksheet
    Dim wsHeatMap As Worksheet
    Dim wsResults As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim opCode As Variant
    Dim testedAVL As Double
    Dim drivP1 As String, respP1 As String
    Dim drivTarget As Double, drivTested As Double
    Dim respTarget As Double, respTested As Double
    Dim drivBenchDiff As Double, respBenchDiff As Double
    Dim drivStatus As String, respStatus As String, finalStatus As String
    Dim outRow As Long
    
    ' NEW: Car selection variables
    Dim targetCarName As String, testedCarName As String
    Dim targetCol As Integer, testedCol As Integer
    Dim cols As Variant
    
    ' Activate Sheet1 so user can see data when selecting cars
    On Error Resume Next
    ThisWorkbook.Sheets("Sheet1").Activate
    On Error GoTo 0

    ' NEW: Show car selection dialog
    If Not ShowCarSelectionDialog() Then
        MsgBox "Evaluation cancelled by user.", vbInformation, "Cancelled"
        Exit Sub
    End If
    
    ' NEW: Get selected car names
    targetCarName = GetSelectedTargetCar()
    testedCarName = GetSelectedTestedCar()
    
    ' NEW: Get column indices for selected cars
    cols = GetSelectedCarColumns()
    targetCol = cols(0)
    testedCol = cols(1)
    
    If targetCol = 0 Or testedCol = 0 Then
        MsgBox "Error: Could not find data columns for selected cars.", vbCritical, "Error"
        Exit Sub
    End If

    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")

    ' Delete existing results sheet if present
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Evaluation Results").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create results sheet
    Set wsResults = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
    wsResults.Name = "Evaluation Results"

    ' Header row with car names
    wsResults.Range("A1:L1").value = Array( _
        "Op Code", "Operation", "Tested AVL", _
        "Driv P1", "Driv Target (" & targetCarName & ")", "Driv Tested (" & testedCarName & ")", "Driv Status", _
        "Resp P1", "Resp Target (" & targetCarName & ")", "Resp Tested (" & testedCarName & ")", "Resp Status", "Final Status")

    With wsResults.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
    End With

    lastRow = wsSheet1.Cells(wsSheet1.Rows.Count, 1).End(xlUp).row
    outRow = 2

    For i = 5 To lastRow
        opCode = wsSheet1.Cells(i, 1).value
        If Trim(CStr(opCode)) <> "" Then
            testedAVL = GetTestedAVL(wsHeatMap, opCode)

            drivP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 5))
            respP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 12))

            ' MODIFIED: Use dynamic columns instead of fixed columns
            drivTarget = ToDbl(wsSheet1.Cells(i, targetCol).value)
            drivTested = ToDbl(wsSheet1.Cells(i, testedCol).value)

            respTarget = ToDbl(wsSheet1.Cells(i, targetCol + 7).value)  ' Resp columns are 7 positions after Driv
            respTested = ToDbl(wsSheet1.Cells(i, testedCol + 6).value)

            drivBenchDiff = benchDiff(drivTarget, drivTested)
            respBenchDiff = benchDiff(respTarget, respTested)

            drivStatus = EvaluateStatus(testedAVL, drivP1, drivBenchDiff, drivTarget, drivTested)
            respStatus = EvaluateStatus(testedAVL, respP1, respBenchDiff, respTarget, respTested)
            finalStatus = CombineStatus(drivStatus, respStatus)

            wsResults.Cells(outRow, 1).value = opCode
            wsResults.Cells(outRow, 2).value = wsSheet1.Cells(i, 2).value
            wsResults.Cells(outRow, 3).value = testedAVL
            wsResults.Cells(outRow, 4).value = drivP1
            wsResults.Cells(outRow, 5).value = drivTarget
            wsResults.Cells(outRow, 6).value = drivTested
            wsResults.Cells(outRow, 7).value = drivStatus
            wsResults.Cells(outRow, 8).value = respP1
            wsResults.Cells(outRow, 9).value = respTarget
            wsResults.Cells(outRow, 10).value = respTested
            wsResults.Cells(outRow, 11).value = respStatus
            wsResults.Cells(outRow, 12).value = finalStatus

            ColorCell wsResults.Cells(outRow, 7), drivStatus
            ColorCell wsResults.Cells(outRow, 11), respStatus
            ColorCell wsResults.Cells(outRow, 12), finalStatus

            outRow = outRow + 1
        End If
    Next i

    wsResults.Columns("A:L").AutoFit

    BuildOperationModeSummary wsResults
    BuildUniqueOverallStatus wsResults

    MsgBox "Evaluation complete!" & vbCrLf & vbCrLf & _
           "Target: " & targetCarName & vbCrLf & _
           "Tested: " & testedCarName & vbCrLf & vbCrLf & _
           "Results written to sheet: " & wsResults.Name, vbInformation, "Success"
End Sub

' [Rest of the functions remain the same as Evaluation_FIXED.bas]
' Builds overall status by op code (unique codes)
Private Sub BuildUniqueOverallStatus(wsResults As Worksheet)
    Dim lastRowRes As Long, i As Long
    Dim code As String, status As String
    Dim dict As Object, nameDict As Object
    Dim arr As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    Set nameDict = CreateObject("Scripting.Dictionary")

    lastRowRes = wsResults.Cells(wsResults.Rows.Count, 1).End(xlUp).row

    For i = 2 To lastRowRes
        code = Trim(CStr(wsResults.Cells(i, 1).value))
        If code <> "" Then
            status = Trim(CStr(wsResults.Cells(i, 12).value))
            If Not dict.Exists(code) Then
                ReDim arr(0 To 0)
                arr(0) = status
                dict.Add code, arr
                nameDict(code) = Trim(CStr(wsResults.Cells(i, 2).value))
            Else
                arr = dict(code)
                ReDim Preserve arr(0 To UBound(arr) + 1)
                arr(UBound(arr)) = status
                dict(code) = arr
            End If
        End If
    Next i

    Dim startRow As Long
    startRow = lastRowRes + 2

    wsResults.Cells(startRow, 1).value = "Overall Status by Op Code"
    With wsResults.Range(wsResults.Cells(startRow, 1), wsResults.Cells(startRow, 4))
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With

    wsResults.Cells(startRow + 1, 1).value = "Op Code"
    wsResults.Cells(startRow + 1, 2).value = "Operation"
    wsResults.Cells(startRow + 1, 3).value = "Overall Status"
    wsResults.Range(wsResults.Cells(startRow + 1, 1), wsResults.Cells(startRow + 1, 3)).Font.Bold = True

    Dim r As Long: r = startRow + 2
    Dim k As Variant

    For Each k In dict.Keys
        Dim anyRed As Boolean: anyRed = False
        Dim allGreen As Boolean: allGreen = True
        arr = dict(k)

        Dim j As Long
        For j = LBound(arr) To UBound(arr)
            status = Trim(CStr(arr(j)))
            If status = "RED" Then anyRed = True
            If status <> "GREEN" Then allGreen = False
        Next j

        Dim overall As String
        If anyRed Then
            overall = "RED"
        ElseIf allGreen Then
            overall = "GREEN"
        Else
            overall = "YELLOW"
        End If

        wsResults.Cells(r, 1).value = k
        wsResults.Cells(r, 2).value = nameDict(k)
        wsResults.Cells(r, 3).value = overall
        ColorCell wsResults.Cells(r, 3), overall
        r = r + 1
    Next k

    wsResults.Columns("A:C").AutoFit
End Sub

' Builds operation mode summary based on a list of known mode codes
Private Sub BuildOperationModeSummary(wsResults As Worksheet)
    Dim modes As Object
    Dim i As Long, lastRowRes As Long
    Dim code As String
    Dim modeStatuses As Object
    Dim mode As Variant
    Dim status As String
    Dim total As Long, yellowCnt As Long
    Dim anyRed As Boolean, allGreen As Boolean
    Dim finalMode As String

    Set modes = CreateObject("Scripting.Dictionary")
    modes.Add "10100000", "Drive away"
    modes.Add "10120000", "Acceleration"
    modes.Add "10030000", "Tip in"
    modes.Add "10040000", "Tip out"
    modes.Add "10070000", "Deceleration"
    modes.Add "10090000", "Gear shift"
    modes.Add "10080000", "Constant speed"
    modes.Add "10010000", "Idle"
    modes.Add "10020000", "Engine start"
    modes.Add "10140000", "Engine shut off"
    modes.Add "10460000", "TCC control"
    modes.Add "10430000", "Cylinder deactivation"
    modes.Add "10450000", "Vehicle stationary"

    lastRowRes = wsResults.Cells(wsResults.Rows.Count, 1).End(xlUp).row
    Set modeStatuses = CreateObject("Scripting.Dictionary")

    For i = 2 To lastRowRes
        code = Trim(CStr(wsResults.Cells(i, 1).value))
        status = Trim(CStr(wsResults.Cells(i, 12).value))

        If modes.Exists(code) Then
            AppendStatus modeStatuses, code, status
        Else
            Dim parent As String
            parent = InferParentMode(code, modes)
            If parent <> "" Then AppendStatus modeStatuses, parent, status
        End If
    Next i

    Dim startRow As Long
    startRow = lastRowRes + 2

    wsResults.Cells(startRow, 6).value = "Operation Mode Summary"
    With wsResults.Range(wsResults.Cells(startRow, 6), wsResults.Cells(startRow, 9))
        .Merge
        .Font.Bold = True
        .Interior.Color = RGB(217, 225, 242)
    End With

    wsResults.Cells(startRow + 1, 6).value = "Op Code"
    wsResults.Cells(startRow + 1, 7).value = "Operation Mode"
    wsResults.Cells(startRow + 1, 8).value = "% Yellow"
    wsResults.Cells(startRow + 1, 9).value = "Final Status"
    wsResults.Range(wsResults.Cells(startRow + 1, 6), wsResults.Cells(startRow + 1, 9)).Font.Bold = True

    Dim r As Long: r = startRow + 2

    For Each mode In modes.Keys
        Dim pctYellow As Double
        total = 0: yellowCnt = 0: anyRed = False: allGreen = True

        If modeStatuses.Exists(mode) Then
            Dim arr As Variant
            arr = modeStatuses(mode)
            Dim j As Long
            For j = LBound(arr) To UBound(arr)
                status = arr(j)
                If status <> "" Then
                    total = total + 1
                    If status = "RED" Then anyRed = True
                    If status = "YELLOW" Then yellowCnt = yellowCnt + 1
                    If status <> "GREEN" Then allGreen = False
                End If
            Next j
        Else
            allGreen = False
        End If

        If total > 0 Then
            pctYellow = yellowCnt / total
        Else
            pctYellow = 0
        End If

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

        wsResults.Cells(r, 6).value = mode
        wsResults.Cells(r, 7).value = modes(mode)
        wsResults.Cells(r, 8).value = pctYellow
        wsResults.Cells(r, 9).value = finalMode
        ColorCell wsResults.Cells(r, 9), finalMode
        r = r + 1
    Next mode

    wsResults.Columns("F:I").AutoFit
End Sub

' Helper: append a status value into a dictionary array
Private Sub AppendStatus(ByRef dict As Object, ByVal key As String, ByVal value As String)
    Dim arr As Variant
    If Not dict.Exists(key) Then
        ReDim arr(0 To 0)
        arr(0) = value
        dict.Add key, arr
    Else
        arr = dict(key)
        ReDim Preserve arr(0 To UBound(arr) + 1)
        arr(UBound(arr)) = value
        dict(key) = arr
    End If
End Sub

' Infer a parent mode code by prefix match
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

' Convert variant to double safely
Private Function ToDbl(v As Variant) As Double
    If IsNumeric(v) Then
        ToDbl = CDbl(v)
    Else
        ToDbl = 0
    End If
End Function

' Bench difference: uses 999 as sentinel when target/tested are both zero or target is zero
Private Function benchDiff(targetVal As Double, testedVal As Double) As Double
    If targetVal = 0 And testedVal = 0 Then
        benchDiff = 999
    ElseIf targetVal = 0 Then
        benchDiff = 999
    Else
        benchDiff = Abs(testedVal - targetVal)
    End If
End Function

' Look up Tested AVL from HeatMap sheet (column 1 = op code; column 8 = AVL)
Private Function GetTestedAVL(wsHeatMap As Worksheet, opCode As Variant) As Double
    Dim opKey As String
    Dim f As Range
    Dim avlCol As Long
    Dim lastRow As Long
    Dim c As Range

    avlCol = 8
    opKey = Trim(CStr(opCode))

    ' Try exact find (string)
    Set f = wsHeatMap.Columns(1).Find(What:=opKey, LookIn:=xlValues, LookAt:=xlWhole, _
        MatchCase:=False, SearchOrder:=xlByRows, SearchDirection:=xlNext)
    If Not f Is Nothing Then
        GetTestedAVL = ToDbl(wsHeatMap.Cells(f.row, avlCol).value)
        Exit Function
    End If

    ' Try numeric match (if opKey numeric-looking)
    If IsNumeric(opKey) Then
        Set f = wsHeatMap.Columns(1).Find(What:=CLng(val(opKey)), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not f Is Nothing Then
            GetTestedAVL = ToDbl(wsHeatMap.Cells(f.row, avlCol).value)
            Exit Function
        End If
    End If

    ' Fall back to manual loop (trimmed string compare)
    lastRow = wsHeatMap.Cells(wsHeatMap.Rows.Count, 1).End(xlUp).row
    For Each c In wsHeatMap.Range(wsHeatMap.Cells(1, 1), wsHeatMap.Cells(lastRow, 1))
        If Trim(CStr(c.value)) = opKey Then
            GetTestedAVL = ToDbl(wsHeatMap.Cells(c.row, avlCol).value)
            Exit Function
        End If
    Next c

    GetTestedAVL = 0
End Function

' Determine P1 status from cell color (prefers DisplayFormat, falls back to Interior/Font)
Private Function GetP1StatusFromColor(rng As Range) As String
    On Error GoTo Fallback
    Dim clr As Long, fclr As Long

    clr = rng.DisplayFormat.Interior.Color
    fclr = rng.DisplayFormat.Font.Color

    GetP1StatusFromColor = MapColorToStatus(clr, fclr)
    If GetP1StatusFromColor <> "N/A" Then Exit Function

Fallback:
    On Error Resume Next
    clr = rng.Interior.Color
    fclr = rng.Font.Color
    GetP1StatusFromColor = MapColorToStatus(clr, fclr)
    If GetP1StatusFromColor = "" Then GetP1StatusFromColor = "N/A"
End Function

' Map fill/font RGB to GREEN / YELLOW / RED / N/A
Private Function MapColorToStatus(fillClr As Long, fontClr As Long) As String
    Dim r As Long, g As Long, b As Long
    Dim rf As Long, gf As Long, bf As Long

    If fillClr > 0 Then
        r = fillClr Mod 256
        g = (fillClr \ 256) Mod 256
        b = (fillClr \ 65536) Mod 256

        If IsNearRGB(r, g, b, 0, 176, 80, 45) Or IsNearRGB(r, g, b, 0, 158, 71, 45) Then
            MapColorToStatus = "GREEN": Exit Function
        End If
        If IsNearRGB(r, g, b, 255, 192, 0, 45) Or IsNearRGB(r, g, b, 255, 217, 102, 60) Then
            MapColorToStatus = "YELLOW": Exit Function
        End If
        If IsNearRGB(r, g, b, 255, 0, 0, 45) Or IsNearRGB(r, g, b, 192, 0, 0, 45) Then
            MapColorToStatus = "RED": Exit Function
        End If
    End If

    If fontClr > 0 Then
        rf = fontClr Mod 256
        gf = (fontClr \ 256) Mod 256
        bf = (fontClr \ 65536) Mod 256

        If IsNearRGB(rf, gf, bf, 0, 128, 0, 5) Then MapColorToStatus = "GREEN": Exit Function
        If IsNearRGB(rf, gf, bf, 255, 255, 0, 5) Then MapColorToStatus = "YELLOW": Exit Function
        If IsNearRGB(rf, gf, bf, 0, 176, 80, 45) Or IsNearRGB(rf, gf, bf, 0, 158, 71, 45) Then MapColorToStatus = "GREEN": Exit Function
        If IsNearRGB(rf, gf, bf, 255, 192, 0, 45) Or IsNearRGB(rf, gf, bf, 255, 217, 102, 60) Then MapColorToStatus = "YELLOW": Exit Function
        If IsNearRGB(rf, gf, bf, 255, 0, 0, 45) Or IsNearRGB(rf, gf, bf, 192, 0, 0, 45) Then MapColorToStatus = "RED": Exit Function
    End If

    MapColorToStatus = "N/A"
End Function

' RGB proximity test
Private Function IsNearRGB(r As Long, g As Long, b As Long, rt As Long, gt As Long, bt As Long, tol As Long) As Boolean
    IsNearRGB = (Abs(r - rt) <= tol) And (Abs(g - gt) <= tol) And (Abs(b - bt) <= tol)
End Function

' Evaluate status using AVL, P1 color and bench difference
Private Function EvaluateStatus(avl As Double, p1 As String, benchDiff As Double, targetVal As Double, testedVal As Double) As String
    ' Updated evaluation logic per specification:
    ' 1. AVL >= 7 AND P1 = GREEN AND meeting benchmark → GREEN (OK)
    ' 2. AVL >= 7 AND P1 = GREEN AND NOT meeting benchmark → YELLOW (Acceptable, improve if possible)
    ' 3. AVL >= 7 AND P1 = YELLOW AND meeting benchmark → YELLOW (Acceptable, improve if possible)
    ' 4. AVL >= 7 AND P1 = YELLOW AND NOT meeting benchmark → YELLOW (Acceptable, improve if possible)
    ' 5. AVL < 7 OR P1 = RED → RED (NOK improve or buy off)
    ' 6. AVL < 7 OR P1 = RED AND meeting benchmark → RED (still NOK improve or buy off)
    ' 7. If no benchmark data → ignore benchmark and evaluate on AVL and P1 only

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

    ' At this point: AVL >= 7 AND P1 is either GREEN or YELLOW
    
    ' Rule 3 & 4: If P1 = YELLOW → Always YELLOW (regardless of benchmark)
    If avl >= 7 And UCase(Trim(p1)) = "YELLOW" Then
        EvaluateStatus = "YELLOW"
        Exit Function
    End If

    ' At this point: AVL >= 7 AND P1 = GREEN
    ' Need to check benchmark data
    
    ' If benchmark data is missing, ignore it and evaluate on AVL/P1 only
    If benchDiff = 999 Then
        ' AVL >= 7 AND P1 = GREEN AND no benchmark data → GREEN
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' If benchmark values not numeric, ignore benchmark
    If Not IsNumeric(targetVal) Or Not IsNumeric(testedVal) Then
        ' AVL >= 7 AND P1 = GREEN AND no valid benchmark → GREEN
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' Benchmark data is available, evaluate it
    ' Rule 1: AVL >= 7 AND P1 = GREEN AND meeting benchmark → GREEN
    ' Rule 2: AVL >= 7 AND P1 = GREEN AND NOT meeting benchmark → YELLOW
    
    ' Meeting benchmark logic:
    ' - If tested >= target → GREEN (always, regardless of how much it exceeds)
    ' - If tested < target AND (target - tested) <= 2 → GREEN (within tolerance)
    ' - If tested < target AND (target - tested) > 2 → YELLOW (not meeting)
    If testedVal >= targetVal Then
        ' Tested meets or exceeds target - always GREEN
        EvaluateStatus = "GREEN"
    Else
        ' Tested below target - check if within tolerance
        If (targetVal - testedVal) <= 2 Then
            ' Within tolerance of 2 units below target
            EvaluateStatus = "GREEN"
        Else
            ' More than 2 units below target - not meeting benchmark
            EvaluateStatus = "YELLOW"
        End If
    End If
End Function


' Combine drive & response statuses into final
' New logic: If one is GREEN and the other is N/A, result is GREEN
' Only show N/A when BOTH are N/A
Private Function CombineStatus(drivStatus As String, respStatus As String) As String
    Dim driv As String, resp As String
    driv = UCase$(Trim$(drivStatus))
    resp = UCase$(Trim$(respStatus))
    
    ' Priority 1: RED - if either is RED, result is RED
    If driv = "RED" Or resp = "RED" Then
        CombineStatus = "RED"
    ' Priority 2: YELLOW - if either is YELLOW (and neither is RED), result is YELLOW
    ElseIf driv = "YELLOW" Or resp = "YELLOW" Then
        CombineStatus = "YELLOW"
    ' Priority 3: GREEN - if at least one is GREEN, result is GREEN
    ElseIf driv = "GREEN" Or resp = "GREEN" Then
        CombineStatus = "GREEN"
    ' Priority 4: N/A - only when BOTH are N/A or blank
    Else
        CombineStatus = "N/A"
    End If
End Function

' Color cell based on status string
Private Sub ColorCell(c As Range, s As String)
    Select Case UCase$(s)
        Case "GREEN"
            c.Interior.Color = RGB(0, 176, 80)
            c.Font.Color = vbWhite
        Case "YELLOW"
            c.Interior.Color = RGB(255, 192, 0)
            c.Font.Color = vbBlack
        Case "RED"
            c.Interior.Color = RGB(192, 0, 0)
            c.Font.Color = vbWhite
        Case Else
            c.Interior.ColorIndex = xlNone
            c.Font.Color = vbBlack
    End Select
End Sub


