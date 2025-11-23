Attribute VB_Name = "HeatMapUpdate_V2"
' ====================================================================
' Module: HeatMapUpdate_V2
' Purpose: Transfer evaluation results to HeatMap Sheet status column
' Version: 2.0 - Enhanced with comprehensive debugging
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
    
    ' Variables for finding sections
    Dim overallStatusRow As Long
    Dim summaryRow As Long
    Dim statusColHeatMap As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugMsg = ""
    
    ' ===== STEP 1: Get worksheets =====
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        ' Try alternative names
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
        If wsHeatMap Is Nothing Then
            Set wsHeatMap = ThisWorkbook.Sheets("Heat Map Sheet")
            If wsHeatMap Is Nothing Then
                MsgBox "ERROR: HeatMap sheet not found!" & vbCrLf & vbCrLf & _
                       "Available sheets: " & ListAllSheets() & vbCrLf & vbCrLf & _
                       "Looking for: 'HeatMap Sheet', 'HeatMap', or 'Heat Map Sheet'", _
                       vbCritical, "Sheet Not Found"
                Exit Sub
            End If
        End If
    End If
    On Error GoTo ErrorHandler
    
    debugMsg = "Sheets found: '" & wsEval.Name & "' and '" & wsHeatMap.Name & "'" & vbCrLf
    
    ' ===== STEP 2: Find "Overall Status by Op Code" section in Evaluation Results =====
    overallStatusRow = FindSectionRow(wsEval, "Overall Status by Op Code")
    If overallStatusRow = 0 Then
        MsgBox "ERROR: Could not find 'Overall Status by Op Code' section in Evaluation Results!" & vbCrLf & vbCrLf & _
               "Diagnostic Info:" & vbCrLf & _
               "- First 20 rows scanned" & vbCrLf & _
               "- Looking for row containing 'Overall Status by Op Code'" & vbCrLf & vbCrLf & _
               "Sample data from rows 1-10:" & vbCrLf & GetSampleData(wsEval, 1, 10), _
               vbCritical, "Section Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & "Found 'Overall Status by Op Code' at row " & overallStatusRow & vbCrLf
    
    ' ===== STEP 3: Find "Operation Mode Summary" section =====
    summaryRow = FindSectionRow(wsEval, "Operation Mode Summary")
    If summaryRow = 0 Then
        debugMsg = debugMsg & "Note: 'Operation Mode Summary' section not found (optional)" & vbCrLf
    Else
        debugMsg = debugMsg & "Found 'Operation Mode Summary' at row " & summaryRow & vbCrLf
    End If
    
    ' ===== STEP 4: Find Status column in HeatMap Sheet =====
    statusColHeatMap = FindStatusColumn(wsHeatMap)
    If statusColHeatMap = 0 Then
        MsgBox "ERROR: Could not find 'Status' column in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Headers found in row 1:" & vbCrLf & GetRowHeaders(wsHeatMap, 1) & vbCrLf & vbCrLf & _
               "Looking for column named 'Status' or 'Current Status'", _
               vbCritical, "Column Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & "Found Status column at column " & statusColHeatMap & " (" & ColumnLetter(statusColHeatMap) & ")" & vbCrLf & vbCrLf
    
    ' ===== STEP 5: Process Overall Status by Op Code section =====
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = debugMsg & "Processing Overall Status by Op Code section..." & vbCrLf
    debugMsg = debugMsg & "- Data rows in Evaluation: " & lastRowEval & vbCrLf
    debugMsg = debugMsg & "- Data rows in HeatMap: " & lastRowHeatMap & vbCrLf & vbCrLf
    
    ' Start from row after "Overall Status by Op Code" header
    For i = overallStatusRow + 2 To lastRowEval  ' +2 to skip header and column titles
        opCode = Trim(wsEval.Cells(i, 1).Value)  ' Column A: Op Code
        finalStatus = Trim(wsEval.Cells(i, 3).Value)  ' Column C: Overall Status
        
        ' Stop if we hit another section or empty rows
        If opCode = "" Or InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
            Exit For
        End If
        
        ' Skip if it's a section header
        If InStr(1, opCode, "Accelerations", vbTextCompare) > 0 Or _
           InStr(1, opCode, "Decelerations", vbTextCompare) > 0 Or _
           InStr(1, opCode, "Drive away", vbTextCompare) > 0 Then
            ' This is a category header, skip
        ElseIf Len(opCode) = 8 And IsNumeric(opCode) Then
            ' Valid operation code, try to update HeatMap
            If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusColHeatMap, lastRowHeatMap) Then
                updatedCount = updatedCount + 1
            End If
        End If
    Next i
    
    debugMsg = debugMsg & "Updated " & updatedCount & " operations from Overall Status section" & vbCrLf & vbCrLf
    
    ' ===== STEP 6: Process Operation Mode Summary section (if found) =====
    If summaryRow > 0 Then
        debugMsg = debugMsg & "Processing Operation Mode Summary section..." & vbCrLf
        Dim summaryCount As Long
        summaryCount = 0
        
        ' Start from row after "Operation Mode Summary" header
        For i = summaryRow + 2 To lastRowEval
            opCode = Trim(wsEval.Cells(i, 6).Value)  ' Column F: Op Code in summary
            finalStatus = Trim(wsEval.Cells(i, 9).Value)  ' Column I: Final Status in summary
            
            ' Stop if we hit empty rows or next section
            If opCode = "" Then
                Exit For
            End If
            
            If Len(opCode) = 8 And IsNumeric(opCode) Then
                If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusColHeatMap, lastRowHeatMap) Then
                    summaryCount = summaryCount + 1
                    updatedCount = updatedCount + 1
                End If
            End If
        Next i
        
        debugMsg = debugMsg & "Updated " & summaryCount & " operations from Summary section" & vbCrLf & vbCrLf
    End If
    
    ' ===== STEP 7: Show results =====
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    If updatedCount > 0 Then
        MsgBox "SUCCESS!" & vbCrLf & vbCrLf & _
               "HeatMap Status Update Complete" & vbCrLf & vbCrLf & _
               "Statistics:" & vbCrLf & _
               "- Operations updated: " & updatedCount & vbCrLf & _
               "- Time taken: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf & vbCrLf & _
               "Details:" & vbCrLf & debugMsg, _
               vbInformation, "Update Complete"
    Else
        MsgBox "WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               "Diagnostic Information:" & vbCrLf & debugMsg & vbCrLf & _
               "Possible causes:" & vbCrLf & _
               "1. Operation codes don't match between sheets" & vbCrLf & _
               "2. No valid status data in Evaluation Results" & vbCrLf & _
               "3. HeatMap sheet doesn't have matching operation codes", _
               vbExclamation, "No Updates"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl & vbCrLf & vbCrLf & _
           "Debug Info:" & vbCrLf & debugMsg, _
           vbCritical, "Error"
End Sub

' Helper function to find a section by searching for text in column A
Private Function FindSectionRow(ws As Worksheet, sectionName As String) As Long
    Dim i As Long
    Dim cellValue As String
    
    FindSectionRow = 0
    
    ' Search first 100 rows for the section
    For i = 1 To 100
        cellValue = Trim(ws.Cells(i, 1).Value)
        If InStr(1, cellValue, sectionName, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find Status column in HeatMap sheet
Private Function FindStatusColumn(ws As Worksheet) As Long
    Dim i As Long
    Dim headerValue As String
    
    FindStatusColumn = 0
    
    ' Search first row for Status column (check first 50 columns)
    For i = 1 To 50
        headerValue = Trim(ws.Cells(1, i).Value)
        If InStr(1, headerValue, "Status", vbTextCompare) > 0 And _
           InStr(1, headerValue, "Current", vbTextCompare) = 0 Then
            ' Found "Status" but not "Current Status" (want the simple Status column)
            FindStatusColumn = i
            Exit Function
        End If
    Next i
    
    ' If not found, look for any column with "Status"
    For i = 1 To 50
        headerValue = Trim(ws.Cells(1, i).Value)
        If UCase(headerValue) = "STATUS" Then
            FindStatusColumn = i
            Exit Function
        End If
    Next i
End Function

' Helper function to update a single row in HeatMap
Private Function UpdateHeatMapRow(ws As Worksheet, opCode As String, status As String, statusCol As Long, lastRow As Long) As Boolean
    Dim i As Long
    Dim heatMapCode As String
    Dim statusDot As String
    Dim statusColor As Long
    
    UpdateHeatMapRow = False
    
    ' Search for matching operation code in HeatMap sheet
    For i = 2 To lastRow  ' Start from row 2 to skip header
        heatMapCode = Trim(ws.Cells(i, 1).Value)
        
        If heatMapCode = opCode Then
            ' Found matching code, update status
            Select Case UCase(Trim(status))
                Case "RED"
                    statusDot = "●"
                    statusColor = RGB(255, 0, 0)  ' Red
                Case "YELLOW"
                    statusDot = "●"
                    statusColor = RGB(255, 255, 0)  ' Yellow
                Case "GREEN"
                    statusDot = "●"
                    statusColor = RGB(0, 255, 0)  ' Green
                Case Else
                    statusDot = "●"
                    statusColor = RGB(128, 128, 128)  ' Gray for N/A
            End Select
            
            ' Update cell
            ws.Cells(i, statusCol).Value = statusDot
            ws.Cells(i, statusCol).Font.Color = statusColor
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Size = 14
            ws.Cells(i, statusCol).HorizontalAlignment = xlCenter
            
            UpdateHeatMapRow = True
            Exit Function
        End If
    Next i
End Function

' Helper function to list all sheet names
Private Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        If sheetList <> "" Then sheetList = sheetList & ", "
        sheetList = sheetList & "'" & ws.Name & "'"
    Next ws
    
    ListAllSheets = sheetList
End Function

' Helper function to get sample data from a range
Private Function GetSampleData(ws As Worksheet, startRow As Long, endRow As Long) As String
    Dim i As Long
    Dim sampleData As String
    
    sampleData = ""
    For i = startRow To endRow
        If ws.Cells(i, 1).Value <> "" Then
            sampleData = sampleData & "Row " & i & ": " & ws.Cells(i, 1).Value & vbCrLf
        End If
    Next i
    
    GetSampleData = sampleData
End Function

' Helper function to get headers from a row
Private Function GetRowHeaders(ws As Worksheet, rowNum As Long) As String
    Dim i As Long
    Dim headers As String
    
    headers = ""
    For i = 1 To 20
        If ws.Cells(rowNum, i).Value <> "" Then
            If headers <> "" Then headers = headers & ", "
            headers = headers & """" & ws.Cells(rowNum, i).Value & """ (Col " & ColumnLetter(i) & ")"
        End If
    Next i
    
    GetRowHeaders = headers
End Function

' Helper function to convert column number to letter
Private Function ColumnLetter(colNum As Long) As String
    Dim d As Long
    Dim m As Long
    Dim name As String
    
    d = colNum
    name = ""
    Do While d > 0
        m = (d - 1) Mod 26
        name = Chr(65 + m) & name
        d = Int((d - m) / 26)
    Loop
    ColumnLetter = name
End Function

' Sub to create Update button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnExists As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("HeatMap")
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets("Heat Map Sheet")
        End If
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Check if button already exists
    On Error Resume Next
    btnExists = Not ws.Buttons("UpdateHeatMapBtn") Is Nothing
    On Error GoTo 0
    
    If btnExists Then
        MsgBox "Button already exists on " & ws.Name & "!", vbInformation
        Exit Sub
    End If
    
    ' Create button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    btn.Name = "UpdateHeatMapBtn"
    btn.Text = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on " & ws.Name & "!" & vbCrLf & vbCrLf & _
           "Click 'Update HeatMap Status' to transfer evaluation results.", _
           vbInformation, "Button Created"
End Sub
