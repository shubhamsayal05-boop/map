Attribute VB_Name = "HeatMapUpdateFixed"
' ====================================================================
' Module: HeatMapUpdateFixed
' Purpose: Transfer evaluation results to HeatMap Sheet status column
' Version: Final Fixed with Enhanced Debugging
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
    Dim debugMsg As String
    Dim statusCol As Long
    Dim foundCount As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    foundCount = 0
    debugMsg = ""
    
    ' Step 1: Get worksheets
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & ListSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")  ' Try alternate name
    End If
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & ListSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing sheets..."
    
    ' Step 2: Find last rows
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = "=== SHEET ANALYSIS ===" & vbCrLf & _
               "Evaluation Results: " & lastRowEval & " rows" & vbCrLf & _
               "HeatMap Sheet: " & lastRowHeatMap & " rows" & vbCrLf & vbCrLf
    
    ' Step 3: Find status column in HeatMap Sheet
    statusCol = FindHeatMapStatusColumn(wsHeatMap)
    
    If statusCol = 0 Then
        debugMsg = debugMsg & "ERROR: Could not find 'Status' column in HeatMap Sheet!" & vbCrLf & _
                   "Headers in row 1: " & GetRowHeaders(wsHeatMap, 1, 10) & vbCrLf
        MsgBox debugMsg, vbCritical, "Column Not Found"
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If
    
    debugMsg = debugMsg & "Status column found at: Column " & ColumnLetter(statusCol) & " (col " & statusCol & ")" & vbCrLf & vbCrLf
    
    ' Step 4: Process "Overall Status by Op Code" section
    Application.StatusBar = "Processing Overall Status by Op Code..."
    Dim overallStart As Long
    overallStart = FindSectionRow(wsEval, "Overall Status by Op Code")
    
    If overallStart > 0 Then
        debugMsg = debugMsg & "=== OVERALL STATUS SECTION ===" & vbCrLf & _
                   "Found at row: " & overallStart & vbCrLf
        
        ' Find "Overall Status" column in this section
        Dim overallStatusCol As Long
        overallStatusCol = FindColumnInRow(wsEval, overallStart + 1, "Overall Status")
        
        If overallStatusCol > 0 Then
            debugMsg = debugMsg & "Overall Status column: " & ColumnLetter(overallStatusCol) & vbCrLf
            
            ' Process rows in this section
            For i = overallStart + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                ' Stop at next section
                If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Or opCode = "" Then
                    Exit For
                End If
                
                If IsNumeric(opCode) And Len(opCode) >= 8 Then
                    foundCount = foundCount + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, overallStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
            
            debugMsg = debugMsg & "Operations found: " & foundCount & vbCrLf & _
                       "Operations updated: " & updatedCount & vbCrLf & vbCrLf
        Else
            debugMsg = debugMsg & "ERROR: 'Overall Status' column not found in section!" & vbCrLf & _
                       "Headers: " & GetRowHeaders(wsEval, overallStart + 1, 10) & vbCrLf & vbCrLf
        End If
    Else
        debugMsg = debugMsg & "WARNING: 'Overall Status by Op Code' section not found!" & vbCrLf & vbCrLf
    End If
    
    ' Step 5: Process "Operation Mode Summary" section
    Application.StatusBar = "Processing Operation Mode Summary..."
    Dim summaryStart As Long
    summaryStart = FindSectionRow(wsEval, "Operation Mode Summary")
    
    If summaryStart > 0 Then
        debugMsg = debugMsg & "=== OPERATION MODE SUMMARY ===" & vbCrLf & _
                   "Found at row: " & summaryStart & vbCrLf
        
        ' Find "Final Status" column in this section
        Dim finalStatusCol As Long
        finalStatusCol = FindColumnInRow(wsEval, summaryStart + 1, "Final Status")
        
        If finalStatusCol > 0 Then
            debugMsg = debugMsg & "Final Status column: " & ColumnLetter(finalStatusCol) & vbCrLf
            
            ' Process rows in this section
            For i = summaryStart + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, 6).Value))  ' Op Code might be in column F
                
                If opCode = "" Then
                    opCode = Trim(CStr(wsEval.Cells(i, 1).Value))  ' Try column A
                End If
                
                ' Stop at end or next section
                If opCode = "" Then
                    Exit For
                End If
                
                If IsNumeric(opCode) And Len(opCode) >= 8 Then
                    foundCount = foundCount + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, finalStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
            
            debugMsg = debugMsg & "Total operations updated: " & updatedCount & vbCrLf
        Else
            debugMsg = debugMsg & "ERROR: 'Final Status' column not found in section!" & vbCrLf & _
                       "Headers: " & GetRowHeaders(wsEval, summaryStart + 1, 10) & vbCrLf
        End If
    Else
        debugMsg = debugMsg & "WARNING: 'Operation Mode Summary' section not found!" & vbCrLf
    End If
    
    ' Step 6: Show results
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 2)
    
    debugMsg = debugMsg & vbCrLf & "=== RESULTS ===" & vbCrLf & _
               "Operations processed: " & foundCount & vbCrLf & _
               "Statuses updated: " & updatedCount & vbCrLf & _
               "Time elapsed: " & elapsed & " seconds"
    
    MsgBox debugMsg, vbInformation, "HeatMap Update Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error updating HeatMap:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           debugMsg, vbCritical, "Error"
End Sub

' Helper: Find section row by title
Private Function FindSectionRow(ws As Worksheet, sectionTitle As String) As Long
    Dim i As Long
    Dim lastRow As Long
    Dim cellValue As String
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    FindSectionRow = 0
    
    For i = 1 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 1).Value))
        If InStr(1, cellValue, sectionTitle, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
End Function

' Helper: Find column in a specific row by header name
Private Function FindColumnInRow(ws As Worksheet, rowNum As Long, headerName As String) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindColumnInRow = 0
    
    For col = 1 To 50  ' Check first 50 columns
        cellValue = Trim(UCase(CStr(ws.Cells(rowNum, col).Value)))
        If InStr(1, cellValue, UCase(headerName), vbTextCompare) > 0 Then
            FindColumnInRow = col
            Exit Function
        End If
    Next col
End Function

' Helper: Find Status column in HeatMap Sheet
Private Function FindHeatMapStatusColumn(ws As Worksheet) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindHeatMapStatusColumn = 0
    
    ' Check row 1 for "Status" header (case insensitive)
    For col = 1 To 50
        cellValue = Trim(UCase(CStr(ws.Cells(1, col).Value)))
        If cellValue = "STATUS" Or InStr(1, cellValue, "CURRENT STATUS", vbTextCompare) > 0 Then
            FindHeatMapStatusColumn = col
            Exit Function
        End If
    Next col
End Function

' Helper: Update a row in HeatMap Sheet
Private Function UpdateHeatMapRow(ws As Worksheet, opCode As String, statusValue As String, statusCol As Long, lastRow As Long) As Boolean
    Dim i As Long
    Dim heatMapOpCode As String
    
    UpdateHeatMapRow = False
    
    ' Search for matching Op Code in HeatMap
    For i = 2 To lastRow  ' Start from row 2 (skip header)
        heatMapOpCode = Trim(CStr(ws.Cells(i, 1).Value))  ' Column A
        
        If heatMapOpCode = opCode Then
            ' Found matching operation - update status
            ws.Cells(i, statusCol).Value = GetStatusSymbol(statusValue)
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Size = 14
            ws.Cells(i, statusCol).Font.Color = GetStatusColor(statusValue)
            ws.Cells(i, statusCol).HorizontalAlignment = xlCenter
            UpdateHeatMapRow = True
            Exit Function
        End If
    Next i
End Function

' Helper: Get status symbol (filled circle)
Private Function GetStatusSymbol(status As String) As String
    GetStatusSymbol = "l"  ' Wingdings filled circle
End Function

' Helper: Get status color
Private Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0)     ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 192, 0)   ' Orange/Yellow
        Case "GREEN"
            GetStatusColor = RGB(0, 176, 80)    ' Green
        Case Else
            GetStatusColor = RGB(128, 128, 128) ' Gray for N/A
    End Select
End Function

' Helper: Get column letter from number
Private Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

' Helper: List all sheet names
Private Function ListSheetNames() As String
    Dim ws As Worksheet
    Dim names As String
    
    names = ""
    For Each ws In ThisWorkbook.Worksheets
        If names <> "" Then names = names & ", "
        names = names & ws.Name
    Next ws
    
    ListSheetNames = names
End Function

' Helper: Get headers from a row
Private Function GetRowHeaders(ws As Worksheet, rowNum As Long, maxCols As Long) As String
    Dim col As Long
    Dim headers As String
    Dim cellValue As String
    
    headers = ""
    For col = 1 To maxCols
        cellValue = Trim(CStr(ws.Cells(rowNum, col).Value))
        If cellValue <> "" Then
            If headers <> "" Then headers = headers & ", "
            headers = headers & ColumnLetter(col) & "=" & cellValue
        End If
    Next col
    
    GetRowHeaders = headers
End Function

' Helper: Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("HeatMap")
    End If
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if present
    btnName = "btnUpdateHeatMap"
    ws.Buttons(btnName).Delete
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    With btn
        .Name = btnName
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
        .Font.Size = 10
        .Font.Bold = True
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update statuses after running evaluation.", _
           vbInformation, "Button Created"
End Sub
