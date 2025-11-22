Attribute VB_Name = "Evaluation"
Option Explicit

' ===============================================
' MODIFIED EVALUATION MODULE WITH CAR SELECTION
' ===============================================
' This version dynamically uses Target and Tested cars selected
' by the user from dropdowns in columns W and X of Sheet1

' Main entry: builds "Evaluation Results" sheet and summaries
' NOW USES SELECTED CARS FROM DROPDOWNS
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
    
    ' CAR SELECTION VARIABLES
    Dim targetCol As Long, testedCol As Long
    Dim targetCar As String, testedCar As String
    
    ' Validate car selections first
    If Not CarSelection.ValidateCarSelections() Then
        Exit Sub
    End If
    
    ' Get selected car columns
    Call CarSelection.GetSelectedCarColumns(targetCol, testedCol)
    If targetCol = 0 Or testedCol = 0 Then
        MsgBox "Error: Could not locate selected car columns. Please check selections.", vbCritical
        Exit Sub
    End If
    
    targetCar = CarSelection.GetTargetCarName()
    testedCar = CarSelection.GetTestedCarName()
    
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
    wsResults.Range("A1").Value = "Op Code"
    wsResults.Range("B1").Value = "Operation"
    wsResults.Range("C1").Value = "Tested AVL"
    wsResults.Range("D1").Value = "Driv P1"
    wsResults.Range("E1").Value = "Driv Target (" & targetCar & ")"
    wsResults.Range("F1").Value = "Driv Tested (" & testedCar & ")"
    wsResults.Range("G1").Value = "Driv Status"
    wsResults.Range("H1").Value = "Resp P1"
    wsResults.Range("I1").Value = "Resp Target (" & targetCar & ")"
    wsResults.Range("J1").Value = "Resp Tested (" & testedCar & ")"
    wsResults.Range("K1").Value = "Resp Status"
    wsResults.Range("L1").Value = "Final Status"

    With wsResults.Range("A1:L1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
        .WrapText = True
    End With

    lastRow = wsSheet1.Cells(wsSheet1.Rows.Count, 1).End(xlUp).Row
    outRow = 2

    For i = 5 To lastRow
        opCode = wsSheet1.Cells(i, 1).Value
        If Trim(CStr(opCode)) <> "" Then
            testedAVL = GetTestedAVL(wsHeatMap, opCode)

            ' Read P1 status from columns E (Driv) and L (Resp) - these are fixed
            drivP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 5))
            respP1 = GetP1StatusFromColor(wsSheet1.Cells(i, 12))

            ' Read Target and Tested values from SELECTED CAR COLUMNS
            ' Target data comes from targetCol for this operation row
            ' Tested data comes from testedCol for this operation row
            drivTarget = ToDbl(wsSheet1.Cells(i, targetCol).Value)  ' Dynamic target column
            drivTested = ToDbl(wsSheet1.Cells(i, testedCol).Value)  ' Dynamic tested column

            ' For Responsiveness section, we need to find the Resp columns
            ' Assuming Resp Target and Resp Tested follow same pattern but in different section
            ' You may need to adjust these based on your actual data structure
            respTarget = ToDbl(wsSheet1.Cells(i, targetCol + 7).Value)  ' Adjust offset as needed
            respTested = ToDbl(wsSheet1.Cells(i, testedCol + 6).Value)  ' Adjust offset as needed

            drivBenchDiff = BenchDiff(drivTarget, drivTested)
            respBenchDiff = BenchDiff(respTarget, respTested)

            drivStatus = EvaluateStatus(testedAVL, drivP1, drivBenchDiff, drivTarget, drivTested)
            respStatus = EvaluateStatus(testedAVL, respP1, respBenchDiff, respTarget, respTested)
            finalStatus = CombineStatus(drivStatus, respStatus)

            wsResults.Cells(outRow, 1).Value = opCode
            wsResults.Cells(outRow, 2).Value = wsSheet1.Cells(i, 2).Value
            wsResults.Cells(outRow, 3).Value = testedAVL
            wsResults.Cells(outRow, 4).Value = drivP1
            wsResults.Cells(outRow, 5).Value = drivTarget
            wsResults.Cells(outRow, 6).Value = drivTested
            wsResults.Cells(outRow, 7).Value = drivStatus
            wsResults.Cells(outRow, 8).Value = respP1
            wsResults.Cells(outRow, 9).Value = respTarget
            wsResults.Cells(outRow, 10).Value = respTested
            wsResults.Cells(outRow, 11).Value = respStatus
            wsResults.Cells(outRow, 12).Value = finalStatus

            ColorCell wsResults.Cells(outRow, 7), drivStatus
            ColorCell wsResults.Cells(outRow, 11), respStatus
            ColorCell wsResults.Cells(outRow, 12), finalStatus

            outRow = outRow + 1
        End If
    Next i

    wsResults.Columns("A:L").AutoFit

    BuildOperationModeSummary wsResults
    BuildUniqueOverallStatus wsResults

    MsgBox "Evaluation complete!" & vbCrLf & _
           "Target Car: " & targetCar & vbCrLf & _
           "Tested Car: " & testedCar & vbCrLf & _
           "Results written to sheet: " & wsResults.Name, vbInformation
End Sub

' Builds overall status by op code (unique codes)
Private Sub BuildUniqueOverallStatus(wsResults As Worksheet)
    Dim lastRow As Long, i As Long
    Dim opCode As String, prevCode As String
    Dim finalStatus As String
    Dim outRow As Long
    Dim wsOverall As Worksheet

    lastRow = wsResults.Cells(wsResults.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then Exit Sub

    ' Sort by op code
    With wsResults.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsResults.Range("A2:A" & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange wsResults.Range("A1:L" & lastRow)
        .Header = xlYes
        .Apply
    End With

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Overall Status by Op Code").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOverall = ThisWorkbook.Sheets.Add(After:=wsResults)
    wsOverall.Name = "Overall Status by Op Code"

    wsOverall.Range("A1:C1").Value = Array("Op Code", "Operation", "Overall Status")
    With wsOverall.Range("A1:C1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
    End With

    outRow = 2
    prevCode = ""

    For i = 2 To lastRow
        opCode = CStr(wsResults.Cells(i, 1).Value)
        If opCode <> prevCode Then
            finalStatus = CStr(wsResults.Cells(i, 12).Value)

            wsOverall.Cells(outRow, 1).Value = opCode
            wsOverall.Cells(outRow, 2).Value = wsResults.Cells(i, 2).Value
            wsOverall.Cells(outRow, 3).Value = finalStatus

            ColorCell wsOverall.Cells(outRow, 3), finalStatus

            outRow = outRow + 1
            prevCode = opCode
        End If
    Next i

    wsOverall.Columns("A:C").AutoFit
End Sub

' Builds operation mode summary with proper sub-operation aggregation
Private Sub BuildOperationModeSummary(wsResults As Worksheet)
    Dim lastRow As Long, i As Long
    Dim opCode As String, parentCode As String
    Dim finalStatus As String
    Dim dict As Object
    Dim modeKey As Variant
    Dim wsOpMode As Worksheet
    Dim outRow As Long
    Dim redCnt As Long, yellowCnt As Long, greenCnt As Long, total As Long
    Dim pctYellow As Double
    Dim allGreen As Boolean, anyRed As Boolean
    Dim finalMode As String

    Set dict = CreateObject("Scripting.Dictionary")

    lastRow = wsResults.Cells(wsResults.Rows.Count, 1).End(xlUp).Row

    ' Aggregate statuses by parent operation mode
    For i = 2 To lastRow
        opCode = CStr(wsResults.Cells(i, 1).Value)
        finalStatus = UCase(Trim(CStr(wsResults.Cells(i, 12).Value)))

        parentCode = InferParentMode(opCode)

        If Not dict.exists(parentCode) Then
            dict(parentCode) = Array(0, 0, 0, wsResults.Cells(i, 2).Value)
        End If

        Dim arr As Variant
        arr = dict(parentCode)

        If finalStatus = "RED" Then
            arr(0) = arr(0) + 1
        ElseIf finalStatus = "YELLOW" Then
            arr(1) = arr(1) + 1
        ElseIf finalStatus = "GREEN" Then
            arr(2) = arr(2) + 1
        End If

        dict(parentCode) = arr
    Next i

    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("Operation Mode Summary").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsOpMode = ThisWorkbook.Sheets.Add(After:=wsResults)
    wsOpMode.Name = "Operation Mode Summary"

    wsOpMode.Range("A1:D1").Value = Array("Op Code", "Operation Mode", "% Yellow", "Final Status")
    With wsOpMode.Range("A1:D1")
        .Font.Bold = True
        .Interior.Color = RGB(68, 114, 196)
        .Font.Color = vbWhite
    End With

    outRow = 2

    For Each modeKey In dict.keys
        Dim arr2 As Variant
        arr2 = dict(modeKey)
        redCnt = arr2(0)
        yellowCnt = arr2(1)
        greenCnt = arr2(2)
        total = redCnt + yellowCnt + greenCnt

        anyRed = (redCnt > 0)
        allGreen = (total > 0 And greenCnt = total)
        pctYellow = 0
        If total > 0 Then pctYellow = yellowCnt / CDbl(total)

        ' FIXED STATUS EVALUATION LOGIC
        If anyRed Then
            finalMode = "RED"
        ElseIf pctYellow > 0.35 Then
            finalMode = "YELLOW"
        ElseIf total > 0 And allGreen Then
            finalMode = "GREEN"
        ElseIf total > 0 Then
            ' Has data but not all green - assign YELLOW
            finalMode = "YELLOW"
        Else
            finalMode = "N/A"
        End If

        wsOpMode.Cells(outRow, 1).Value = modeKey
        wsOpMode.Cells(outRow, 2).Value = arr2(3)
        wsOpMode.Cells(outRow, 3).Value = pctYellow
        wsOpMode.Cells(outRow, 4).Value = finalMode

        ColorCell wsOpMode.Cells(outRow, 4), finalMode

        outRow = outRow + 1
    Next modeKey

    wsOpMode.Columns("A:D").AutoFit
    wsOpMode.Columns("C").NumberFormat = "0.00%"
End Sub

' FIXED: Infer parent operation mode from sub-operation code
' Matches on first 4 digits instead of all 8
Private Function InferParentMode(code As String) As String
    Dim dict As Object
    Dim k As Variant

    Set dict = CreateObject("Scripting.Dictionary")
    dict("10100000") = True
    dict("10120000") = True
    dict("10030000") = True
    dict("10040000") = True
    dict("10070000") = True
    dict("10090000") = True
    dict("10080000") = True
    dict("10010000") = True
    dict("10020000") = True
    dict("10140000") = True
    dict("10460000") = True
    dict("10430000") = True
    dict("10450000") = True

    ' Check if code itself is a parent
    If dict.exists(code) Then
        InferParentMode = code
        Exit Function
    End If

    ' FIXED: Match on first 4 digits
    For Each k In dict.keys
        If Len(code) >= 4 And Len(k) >= 4 Then
            If Left$(code, 4) = Left$(k, 4) Then
                InferParentMode = k
                Exit Function
            End If
        End If
    Next k

    InferParentMode = code
End Function

' FIXED: Evaluate status with rule-based logic and tolerance
Private Function EvaluateStatus(avl As Double, p1 As String, _
                                benchDiff As Double, targetVal As Double, testedVal As Double) As String
    p1 = UCase(Trim(p1))

    ' If P1 is N/A or blank, return blank
    If p1 = "N/A" Or p1 = "" Then
        EvaluateStatus = ""
        Exit Function
    End If

    ' Rule 5 & 6: AVL < 7 OR P1 = RED → Always RED
    If avl < 7 Or p1 = "RED" Then
        EvaluateStatus = "RED"
        Exit Function
    End If

    ' Rule 3 & 4: P1 = YELLOW → Always YELLOW
    If avl >= 7 And p1 = "YELLOW" Then
        EvaluateStatus = "YELLOW"
        Exit Function
    End If

    ' At this point: AVL >= 7 AND P1 = GREEN

    ' Rule 7: If benchmark data missing, ignore it
    If benchDiff = 999 Then
        EvaluateStatus = "GREEN"
        Exit Function
    End If

    ' Rule 1 & 2: Benchmark comparison with tolerance
    If testedVal >= targetVal Then
        ' Tested meets or exceeds target - always GREEN
        EvaluateStatus = "GREEN"
    Else
        ' Tested below target - check tolerance (2 units)
        If (targetVal - testedVal) <= 2 Then
            EvaluateStatus = "GREEN"   ' Within tolerance
        Else
            EvaluateStatus = "YELLOW"  ' Not meeting benchmark
        End If
    End If
End Function

' Combine Driv + Resp → Final
Private Function CombineStatus(drivStatus As String, respStatus As String) As String
    drivStatus = UCase(Trim(drivStatus))
    respStatus = UCase(Trim(respStatus))

    If drivStatus = "RED" Or respStatus = "RED" Then
        CombineStatus = "RED"
    ElseIf drivStatus = "YELLOW" Or respStatus = "YELLOW" Then
        CombineStatus = "YELLOW"
    ElseIf drivStatus = "GREEN" And respStatus = "GREEN" Then
        CombineStatus = "GREEN"
    Else
        CombineStatus = ""
    End If
End Function

' Benchmark difference
Private Function BenchDiff(target As Double, tested As Double) As Double
    If target = 0 And tested = 0 Then
        BenchDiff = 999
    Else
        BenchDiff = tested - target
    End If
End Function

' Safe conversion to Double
Private Function ToDbl(val As Variant) As Double
    On Error Resume Next
    ToDbl = CDbl(val)
    If Err.Number <> 0 Then ToDbl = 0
    On Error GoTo 0
End Function

' Get tested AVL from HeatMap
Private Function GetTestedAVL(wsHeatMap As Worksheet, opCode As Variant) As Double
    Dim lastRow As Long, i As Long
    lastRow = wsHeatMap.Cells(wsHeatMap.Rows.Count, 1).End(xlUp).Row

    For i = 2 To lastRow
        If CStr(wsHeatMap.Cells(i, 1).Value) = CStr(opCode) Then
            GetTestedAVL = ToDbl(wsHeatMap.Cells(i, 3).Value)
            Exit Function
        End If
    Next i

    GetTestedAVL = 0
End Function

' Get P1 status from cell color
Private Function GetP1StatusFromColor(cell As Range) As String
    Dim clr As Long
    clr = cell.Interior.Color

    If clr = RGB(255, 0, 0) Or clr = RGB(192, 0, 0) Then
        GetP1StatusFromColor = "RED"
    ElseIf clr = RGB(255, 255, 0) Or clr = RGB(255, 192, 0) Then
        GetP1StatusFromColor = "YELLOW"
    ElseIf clr = RGB(0, 255, 0) Or clr = RGB(0, 176, 80) Or clr = RGB(146, 208, 80) Then
        GetP1StatusFromColor = "GREEN"
    Else
        GetP1StatusFromColor = "N/A"
    End If
End Function

' Color cell based on status
Private Sub ColorCell(cell As Range, status As String)
    status = UCase(Trim(status))

    Select Case status
        Case "RED"
            cell.Interior.Color = RGB(255, 0, 0)
            cell.Font.Color = vbWhite
            cell.Font.Bold = True
        Case "YELLOW"
            cell.Interior.Color = RGB(255, 255, 0)
            cell.Font.Color = vbBlack
            cell.Font.Bold = True
        Case "GREEN"
            cell.Interior.Color = RGB(0, 255, 0)
            cell.Font.Color = vbBlack
            cell.Font.Bold = True
        Case Else
            cell.Interior.ColorIndex = xlNone
            cell.Font.Color = vbBlack
    End Select
End Sub
