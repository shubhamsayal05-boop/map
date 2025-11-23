Attribute VB_Name = "HeatMapUpdate_v2"
' ====================================================================
' Module: HeatMapUpdate_v2
' Purpose: Transfer evaluation results to HeatMap Sheet with enhanced debugging
' Version: 2.0 - Fixed for actual workbook structure
' ====================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results
Sub UpdateHeatMapStatus()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim lastRowEval As Long
    Dim lastRowHeatMap As Long
    Dim i As Long, j As Long
    Dim opCode As String
    Dim finalStatus As String
    Dim updatedCount As Long
    Dim startTime As Double
    Dim debugInfo As String
    Dim statusCol As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugInfo = "=== HEATMAP UPDATE DIAGNOSTIC ===" & vbCrLf & vbCrLf
    
    ' Step 1: Check if sheets exist
    debugInfo = debugInfo & "STEP 1: Checking sheets..." & vbCrLf
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & ListSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ Evaluation Results sheet found" & vbCrLf
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & ListSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ HeatMap Sheet found" & vbCrLf & vbCrLf
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    ' Step 2: Analyze Evaluation Results structure
    debugInfo = debugInfo & "STEP 2: Analyzing Evaluation Results sheet..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Last row with data: " & lastRowEval & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    Dim overallRow As Long, summaryRow As Long
    overallRow = FindTextInColumn(wsEval, "Overall Status by Op Code", 1, lastRowEval)
    summaryRow = FindTextInColumn(wsEval, "Operation Mode Summary", 1, lastRowEval)
    
    If overallRow > 0 Then
        debugInfo = debugInfo & "  ✓ 'Overall Status by Op Code' found at row " & overallRow & vbCrLf
    Else
        debugInfo = debugInfo & "  ✗ 'Overall Status by Op Code' NOT found" & vbCrLf
    End If
    
    If summaryRow > 0 Then
        debugInfo = debugInfo & "  ✓ 'Operation Mode Summary' found at row " & summaryRow & vbCrLf
    Else
        debugInfo = debugInfo & "  ✗ 'Operation Mode Summary' NOT found" & vbCrLf
    End If
    debugInfo = debugInfo & vbCrLf
    
    ' Step 3: Analyze HeatMap Sheet structure
    debugInfo = debugInfo & "STEP 3: Analyzing HeatMap Sheet..." & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Last row with data: " & lastRowHeatMap & vbCrLf
    
    ' Find "Status" column in HeatMap
    statusCol = FindColumnByName(wsHeatMap, "Status", 1)
    If statusCol > 0 Then
        debugInfo = debugInfo & "  ✓ 'Status' column found at column " & statusCol & " (" & ColumnLetter(statusCol) & ")" & vbCrLf
    Else
        ' Try alternative names
        statusCol = FindColumnByName(wsHeatMap, "Current Status", 1)
        If statusCol > 0 Then
            debugInfo = debugInfo & "  ✓ 'Current Status' column found at column " & statusCol & " (" & ColumnLetter(statusCol) & ")" & vbCrLf
        Else
            debugInfo = debugInfo & "  ✗ 'Status' column NOT found" & vbCrLf
            debugInfo = debugInfo & "  First 10 column headers:" & vbCrLf
            For i = 1 To Application.Min(10, wsHeatMap.Cells(1, wsHeatMap.Columns.Count).End(xlToLeft).Column)
                debugInfo = debugInfo & "    Col " & ColumnLetter(i) & ": " & wsHeatMap.Cells(1, i).Value & vbCrLf
            Next i
        End If
    End If
    debugInfo = debugInfo & vbCrLf
    
    ' Step 4: Process data
    debugInfo = debugInfo & "STEP 4: Processing operations..." & vbCrLf
    
    If overallRow > 0 And statusCol > 0 Then
        ' Find the header row (should be next row after section title)
        Dim headerRow As Long
        headerRow = overallRow + 1
        
        ' Find "Final Status" or "Overall Status" column
        Dim finalStatusCol As Long
        finalStatusCol = FindColumnInRow(wsEval, headerRow, "Final Status")
        If finalStatusCol = 0 Then
            finalStatusCol = FindColumnInRow(wsEval, headerRow, "Overall Status")
        End If
        
        If finalStatusCol > 0 Then
            debugInfo = debugInfo & "  ✓ Status column in Evaluation Results: " & finalStatusCol & " (" & ColumnLetter(finalStatusCol) & ")" & vbCrLf
            
            ' Process data rows
            Dim dataStartRow As Long
            dataStartRow = headerRow + 1
            
            Dim opCodesProcessed As Long, opCodesMatched As Long
            opCodesProcessed = 0
            opCodesMatched = 0
            
            ' Loop until we hit next section or end
            For i = dataStartRow To lastRowEval
                ' Check if we hit the next section
                Dim cellValue As String
                cellValue = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                If InStr(1, cellValue, "Operation Mode Summary", vbTextCompare) > 0 Or _
                   InStr(1, cellValue, "Accelerations", vbTextCompare) > 0 Or _
                   InStr(1, cellValue, "Decelerations", vbTextCompare) > 0 Then
                    Exit For
                End If
                
                ' Process if it's a valid Op Code
                If cellValue <> "" And IsNumeric(cellValue) And Len(cellValue) >= 8 Then
                    opCodesProcessed = opCodesProcessed + 1
                    opCode = cellValue
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, finalStatusCol).Value)))
                    
                    ' Update in HeatMap
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        Dim heatMapRow As Long
                        heatMapRow = FindOpCodeInHeatMap(wsHeatMap, opCode, lastRowHeatMap)
                        
                        If heatMapRow > 0 Then
                            ' Update the status
                            wsHeatMap.Cells(heatMapRow, statusCol).Value = GetStatusDot(finalStatus)
                            wsHeatMap.Cells(heatMapRow, statusCol).Font.Name = "Wingdings"
                            wsHeatMap.Cells(heatMapRow, statusCol).Font.Size = 14
                            wsHeatMap.Cells(heatMapRow, statusCol).Font.Color = GetStatusColor(finalStatus)
                            
                            opCodesMatched = opCodesMatched + 1
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
            
            debugInfo = debugInfo & "  Operations in Evaluation Results: " & opCodesProcessed & vbCrLf
            debugInfo = debugInfo & "  Operations matched in HeatMap: " & opCodesMatched & vbCrLf
            debugInfo = debugInfo & "  Status updates applied: " & updatedCount & vbCrLf
        Else
            debugInfo = debugInfo & "  ✗ Could not find 'Final Status' column in Evaluation Results" & vbCrLf
            debugInfo = debugInfo & "  Header row content:" & vbCrLf
            For i = 1 To Application.Min(15, wsEval.Cells(headerRow, wsEval.Columns.Count).End(xlToLeft).Column)
                debugInfo = debugInfo & "    Col " & ColumnLetter(i) & ": " & wsEval.Cells(headerRow, i).Value & vbCrLf
            Next i
        End If
    Else
        debugInfo = debugInfo & "  ✗ Cannot process: Missing section or status column" & vbCrLf
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Complete
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    
    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 2)
    
    debugInfo = debugInfo & "RESULT:" & vbCrLf
    debugInfo = debugInfo & "  Updated operations: " & updatedCount & vbCrLf
    debugInfo = debugInfo & "  Time elapsed: " & elapsed & " seconds" & vbCrLf
    
    ' Show results
    If updatedCount > 0 Then
        MsgBox "HeatMap Status Update Complete!" & vbCrLf & vbCrLf & _
               "Operations updated: " & updatedCount & vbCrLf & _
               "Time elapsed: " & elapsed & " seconds", _
               vbInformation, "Update Complete"
    Else
        ' Show detailed diagnostic
        MsgBox debugInfo, vbExclamation, "No Updates Made - Diagnostic Information"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.StatusBar = False
    MsgBox "Error occurred: " & Err.Description & vbCrLf & vbCrLf & debugInfo, vbCritical, "Error"
End Sub

' Helper function: Find text in a column
Function FindTextInColumn(ws As Worksheet, searchText As String, col As Long, lastRow As Long) As Long
    Dim i As Long
    FindTextInColumn = 0
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, col).Value), searchText, vbTextCompare) > 0 Then
            FindTextInColumn = i
            Exit Function
        End If
    Next i
End Function

' Helper function: Find column by name in row 1
Function FindColumnByName(ws As Worksheet, colName As String, searchRow As Long) As Long
    Dim i As Long
    Dim lastCol As Long
    FindColumnByName = 0
    lastCol = ws.Cells(searchRow, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If InStr(1, CStr(ws.Cells(searchRow, i).Value), colName, vbTextCompare) > 0 Then
            FindColumnByName = i
            Exit Function
        End If
    Next i
End Function

' Helper function: Find column by name in specific row
Function FindColumnInRow(ws As Worksheet, row As Long, colName As String) As Long
    Dim i As Long
    Dim lastCol As Long
    FindColumnInRow = 0
    lastCol = ws.Cells(row, ws.Columns.Count).End(xlToLeft).Column
    For i = 1 To lastCol
        If InStr(1, CStr(ws.Cells(row, i).Value), colName, vbTextCompare) > 0 Then
            FindColumnInRow = i
            Exit Function
        End If
    Next i
End Function

' Helper function: Find Op Code in HeatMap
Function FindOpCodeInHeatMap(ws As Worksheet, opCode As String, lastRow As Long) As Long
    Dim i As Long
    FindOpCodeInHeatMap = 0
    For i = 1 To lastRow
        If Trim(CStr(ws.Cells(i, 1).Value)) = opCode Then
            FindOpCodeInHeatMap = i
            Exit Function
        End If
    Next i
End Function

' Helper function: Get status dot character
Function GetStatusDot(status As String) As String
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusDot = "l"  ' Filled circle in Wingdings
        Case "YELLOW"
            GetStatusDot = "l"
        Case "GREEN"
            GetStatusDot = "l"
        Case Else
            GetStatusDot = "l"  ' Gray for N/A or unknown
    End Select
End Function

' Helper function: Get status color
Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0)     ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 192, 0)   ' Yellow/Orange
        Case "GREEN"
            GetStatusColor = RGB(0, 176, 80)    ' Green
        Case Else
            GetStatusColor = RGB(166, 166, 166) ' Gray
    End Select
End Function

' Helper function: Convert column number to letter
Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

' Helper function: List all sheet names
Function ListSheetNames() As String
    Dim ws As Worksheet
    Dim names As String
    names = ""
    For Each ws In ThisWorkbook.Sheets
        names = names & ws.Name & ", "
    Next ws
    If Len(names) > 2 Then names = Left(names, Len(names) - 2)
    ListSheetNames = names
End Function

' Helper function: Create Update button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbCritical
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Delete existing button if any
    btnName = "btnUpdateHeatMap"
    On Error Resume Next
    ws.Buttons(btnName).Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 180, 30)
    btn.Name = btnName
    btn.Text = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    btn.Font.Bold = True
    btn.Font.Size = 11
    
    MsgBox "Button created successfully!" & vbCrLf & vbCrLf & _
           "Click the 'Update HeatMap Status' button to transfer evaluation results.", _
           vbInformation, "Button Created"
End Sub
