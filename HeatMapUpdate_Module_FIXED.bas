Attribute VB_Name = "HeatMapUpdate"
' ====================================================================
' Module: HeatMapUpdate_FIXED
' Purpose: Transfer evaluation results to HeatMap Sheet status column
' Version: 2.0 - Fixed with comprehensive debugging
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
    Dim statusColIndex As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugInfo = "=== HeatMap Status Update Debug Info ===" & vbCrLf & vbCrLf
    
    ' Step 1: Find worksheets
    debugInfo = debugInfo & "Step 1: Finding worksheets..." & vbCrLf
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "ERROR: Cannot find 'Evaluation Results' sheet!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: Cannot find 'HeatMap Sheet'!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    debugInfo = debugInfo & "  ✓ Found 'Evaluation Results'" & vbCrLf
    debugInfo = debugInfo & "  ✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    
    ' Step 2: Analyze Evaluation Results structure
    debugInfo = debugInfo & "Step 2: Analyzing Evaluation Results sheet..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Total rows in Evaluation Results: " & lastRowEval & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    Dim overallRow As Long
    overallRow = FindTextInColumn(wsEval, 1, "Overall Status by Op Code", lastRowEval)
    
    If overallRow > 0 Then
        debugInfo = debugInfo & "  ✓ Found 'Overall Status by Op Code' at row " & overallRow & vbCrLf
        debugInfo = debugInfo & "    Headers in row " & (overallRow + 1) & ":" & vbCrLf
        debugInfo = debugInfo & "    " & GetRowHeaders(wsEval, overallRow + 1, 5) & vbCrLf
    Else
        debugInfo = debugInfo & "  ✗ 'Overall Status by Op Code' section NOT FOUND!" & vbCrLf
    End If
    
    ' Find "Operation Mode Summary" section
    Dim summaryRow As Long
    summaryRow = FindTextInColumn(wsEval, 1, "Operation Mode Summary", lastRowEval)
    
    If summaryRow > 0 Then
        debugInfo = debugInfo & "  ✓ Found 'Operation Mode Summary' at row " & summaryRow & vbCrLf
        debugInfo = debugInfo & "    Headers in row " & (summaryRow + 1) & ":" & vbCrLf
        debugInfo = debugInfo & "    " & GetRowHeaders(wsEval, summaryRow + 1, 5) & vbCrLf
    Else
        debugInfo = debugInfo & "  ✗ 'Operation Mode Summary' section NOT FOUND!" & vbCrLf
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Step 3: Analyze HeatMap Sheet structure
    debugInfo = debugInfo & "Step 3: Analyzing HeatMap Sheet..." & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Total rows in HeatMap Sheet: " & lastRowHeatMap & vbCrLf
    debugInfo = debugInfo & "  First 5 column headers:" & vbCrLf
    debugInfo = debugInfo & "    " & GetRowHeaders(wsHeatMap, 1, 10) & vbCrLf
    
    ' Find "Status" column in HeatMap
    statusColIndex = FindColumnByHeader(wsHeatMap, 1, "Status")
    
    If statusColIndex > 0 Then
        debugInfo = debugInfo & "  ✓ Found 'Status' column at position " & statusColIndex & _
                   " (" & ColumnLetter(statusColIndex) & ")" & vbCrLf
    Else
        ' Try alternate names
        statusColIndex = FindColumnByHeader(wsHeatMap, 1, "Current Status")
        If statusColIndex = 0 Then statusColIndex = FindColumnByHeader(wsHeatMap, 1, "Status P1")
        If statusColIndex = 0 Then statusColIndex = FindColumnByHeader(wsHeatMap, 1, "Current Status P1")
        
        If statusColIndex > 0 Then
            debugInfo = debugInfo & "  ✓ Found status column at position " & statusColIndex & _
                       " (" & ColumnLetter(statusColIndex) & ")" & vbCrLf
        Else
            debugInfo = debugInfo & "  ✗ Could NOT find 'Status' column!" & vbCrLf
        End If
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Sample operation codes from both sheets
    debugInfo = debugInfo & "Step 4: Sample operation codes..." & vbCrLf
    debugInfo = debugInfo & "  From Evaluation Results (first 3 after Overall Status):" & vbCrLf
    If overallRow > 0 Then
        For i = overallRow + 2 To Application.Min(overallRow + 4, lastRowEval)
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            If opCode <> "" And IsNumeric(opCode) Then
                debugInfo = debugInfo & "    Row " & i & ": " & opCode & vbCrLf
            End If
        Next i
    End If
    
    debugInfo = debugInfo & "  From HeatMap Sheet (first 3 codes):" & vbCrLf
    For i = 2 To Application.Min(4, lastRowHeatMap)
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" Then
            debugInfo = debugInfo & "    Row " & i & ": " & opCode & vbCrLf
        End If
    Next i
    
    debugInfo = debugInfo & vbCrLf
    
    ' Check if we can proceed
    If overallRow = 0 And summaryRow = 0 Then
        MsgBox debugInfo & vbCrLf & _
               "CANNOT PROCEED: No evaluation sections found!" & vbCrLf & vbCrLf & _
               "Please ensure evaluation has been run.", vbCritical, "No Evaluation Data"
        Exit Sub
    End If
    
    If statusColIndex = 0 Then
        MsgBox debugInfo & vbCrLf & _
               "CANNOT PROCEED: Cannot find Status column in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Please ensure HeatMap Sheet has a 'Status' column.", vbCritical, "Status Column Not Found"
        Exit Sub
    End If
    
    ' Step 5: Update statuses
    debugInfo = debugInfo & "Step 5: Updating statuses..." & vbCrLf
    Application.ScreenUpdating = False
    
    ' Process "Overall Status by Op Code" section
    If overallRow > 0 Then
        Dim finalStatusCol As Long
        finalStatusCol = FindColumnByHeader(wsEval, overallRow + 1, "Final Status")
        
        If finalStatusCol = 0 Then finalStatusCol = FindColumnByHeader(wsEval, overallRow + 1, "Overall Status")
        
        If finalStatusCol > 0 Then
            debugInfo = debugInfo & "  Processing Overall Status section (Final Status in column " & _
                       ColumnLetter(finalStatusCol) & ")..." & vbCrLf
            
            For i = overallRow + 2 To lastRowEval
                ' Stop if we hit next section
                If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
                    Exit For
                End If
                
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                If opCode <> "" And IsNumeric(opCode) Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, finalStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                        ' Update in HeatMap
                        If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusColIndex, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
        End If
    End If
    
    ' Process "Operation Mode Summary" section
    If summaryRow > 0 Then
        Dim summaryStatusCol As Long
        summaryStatusCol = FindColumnByHeader(wsEval, summaryRow + 1, "Final Status")
        
        If summaryStatusCol > 0 Then
            debugInfo = debugInfo & "  Processing Operation Mode Summary section (Final Status in column " & _
                       ColumnLetter(summaryStatusCol) & ")..." & vbCrLf
            
            For i = summaryRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                ' Stop if empty rows
                If opCode = "" Then Exit For
                
                If IsNumeric(opCode) Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, summaryStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                        ' Update in HeatMap
                        If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusColIndex, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
        End If
    End If
    
    Application.ScreenUpdating = True
    
    debugInfo = debugInfo & vbCrLf & "=== UPDATE COMPLETE ===" & vbCrLf & _
               "Updated " & updatedCount & " operations in " & Format(Timer - startTime, "0.00") & " seconds"
    
    ' Show results
    If updatedCount > 0 Then
        MsgBox "✓ Successfully updated " & updatedCount & " operations!" & vbCrLf & vbCrLf & _
               "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf & vbCrLf & _
               "Debug info saved to Immediate window (Ctrl+G to view)", _
               vbInformation, "HeatMap Status Updated"
    Else
        MsgBox "⚠ No statuses were updated!" & vbCrLf & vbCrLf & _
               "Please review the debug information below:" & vbCrLf & vbCrLf & debugInfo, _
               vbExclamation, "No Updates Made"
    End If
    
    ' Print debug info to Immediate window
    Debug.Print String(80, "=")
    Debug.Print debugInfo
    Debug.Print String(80, "=")
    
    Application.StatusBar = False
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Debug info:" & vbCrLf & debugInfo, vbCritical, "Error in UpdateHeatMapStatus"
End Sub

' Helper: Update a single row in HeatMap
Function UpdateHeatMapRow(ws As Worksheet, opCode As String, status As String, _
                         statusCol As Long, lastRow As Long) As Boolean
    Dim i As Long
    Dim heatMapCode As String
    
    UpdateHeatMapRow = False
    
    ' Find the operation in HeatMap
    For i = 2 To lastRow ' Start from row 2 (skip header)
        heatMapCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If heatMapCode = opCode Then
            ' Found it - update status with colored dot
            ws.Cells(i, statusCol).Value = GetStatusDot(status)
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Size = 14
            ws.Cells(i, statusCol).Font.Color = GetStatusColor(status)
            ws.Cells(i, statusCol).HorizontalAlignment = xlCenter
            UpdateHeatMapRow = True
            Exit Function
        End If
    Next i
End Function

' Helper: Get colored dot for status
Function GetStatusDot(status As String) As String
    GetStatusDot = "l" ' Wingdings filled circle
End Function

' Helper: Get RGB color for status
Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0) ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 192, 0) ' Orange/Yellow
        Case "GREEN"
            GetStatusColor = RGB(0, 176, 80) ' Green
        Case Else
            GetStatusColor = RGB(128, 128, 128) ' Gray for N/A
    End Select
End Function

' Helper: Find text in column A
Function FindTextInColumn(ws As Worksheet, col As Long, searchText As String, lastRow As Long) As Long
    Dim i As Long
    FindTextInColumn = 0
    
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, col).Value), searchText, vbTextCompare) > 0 Then
            FindTextInColumn = i
            Exit Function
        End If
    Next i
End Function

' Helper: Find column by header name
Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindColumnByHeader = 0
    
    For col = 1 To 50 ' Check first 50 columns
        cellValue = Trim(UCase(CStr(ws.Cells(headerRow, col).Value)))
        If InStr(1, cellValue, UCase(headerName), vbTextCompare) > 0 Then
            FindColumnByHeader = col
            Exit Function
        End If
    Next col
End Function

' Helper: Get row headers for debugging
Function GetRowHeaders(ws As Worksheet, row As Long, numCols As Long) As String
    Dim col As Long
    Dim result As String
    
    result = ""
    For col = 1 To numCols
        If col > 1 Then result = result & " | "
        result = result & ColumnLetter(col) & ": " & Trim(CStr(ws.Cells(row, col).Value))
    Next col
    
    GetRowHeaders = result
End Function

' Helper: Convert column number to letter
Function ColumnLetter(colNum As Long) As String
    Dim temp As Long
    Dim letter As String
    
    temp = colNum
    Do While temp > 0
        temp = temp - 1
        letter = Chr(65 + (temp Mod 26)) & letter
        temp = temp \ 26
    Loop
    
    ColumnLetter = letter
End Function

' Helper: List all sheets in workbook
Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim result As String
    
    result = ""
    For Each ws In ThisWorkbook.Worksheets
        result = result & "  - " & ws.Name & vbCrLf
    Next ws
    
    ListAllSheets = result
End Function

' Create a button on HeatMap Sheet to run the update
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Object
    Dim btnLeft As Double
    Dim btnTop As Double
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Cannot find 'HeatMap Sheet'!" & vbCrLf & vbCrLf & _
               "Please ensure the sheet is named exactly 'HeatMap Sheet'", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Remove existing button if any
    On Error Resume Next
    ws.Buttons("UpdateHeatMapBtn").Delete
    On Error GoTo 0
    
    ' Position for button (top-right area)
    btnLeft = ws.Range("K1").Left
    btnTop = ws.Range("K1").Top
    
    ' Create button
    Set btn = ws.Buttons.Add(btnLeft, btnTop, 180, 30)
    With btn
        .Name = "UpdateHeatMapBtn"
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    MsgBox "✓ Button created successfully!" & vbCrLf & vbCrLf & _
           "Look for 'Update HeatMap Status' button on HeatMap Sheet" & vbCrLf & _
           "Click it after running evaluation to fill status dots", _
           vbInformation, "Button Created"
End Sub
