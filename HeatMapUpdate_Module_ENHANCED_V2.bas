Attribute VB_Name = "HeatMapUpdate_Enhanced_V2"
' ==============================================================================
' HeatMap Status Update Module - Enhanced Version 2 with Full Diagnostics
' ==============================================================================
' This module transfers evaluation results to HeatMap Sheet with detailed debug info
'
' REQUIREMENTS (based on user feedback):
' - Evaluation Results sheet has:
'   * "Overall Status by Op Code" section with codes and statuses
'   * "Operation Mode Summary" section with codes and statuses
' - HeatMap Sheet has:
'   * Column A: Operation codes
'   * Column with header "Status": Where colored dots will be filled
'
' Usage:
'   1. Import this module (Alt+F11 → File → Import)
'   2. Run UpdateHeatMapStatusEnhanced() macro
'   3. Review detailed diagnostic messages
' ==============================================================================

Option Explicit

' Main function to update HeatMap Status with comprehensive diagnostics
Sub UpdateHeatMapStatusEnhanced()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim evalLastRow As Long
    Dim heatMapLastRow As Long
    Dim statusCol As Long
    Dim i As Long, j As Long
    Dim opCode As String
    Dim status As String
    Dim updated As Long
    Dim notFound As Long
    Dim startTime As Double
    Dim diagnostics As String
    
    On Error GoTo ErrorHandler
    startTime = Timer
    updated = 0
    notFound = 0
    diagnostics = ""
    
    ' Step 1: Find and validate Evaluation Results sheet
    diagnostics = diagnostics & "=== STEP 1: Finding Evaluation Results Sheet ===" & vbCrLf
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    diagnostics = diagnostics & "✓ Found 'Evaluation Results' sheet" & vbCrLf & vbCrLf
    
    ' Step 2: Find and validate HeatMap Sheet
    diagnostics = diagnostics & "=== STEP 2: Finding HeatMap Sheet ===" & vbCrLf
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
    End If
    If wsHeatMap Is Nothing Then
        Set wsHeatMap = ThisWorkbook.Sheets("Heat Map")
    End If
    On Error GoTo ErrorHandler
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: HeatMap Sheet not found!" & vbCrLf & vbCrLf & _
               "Tried: 'HeatMap Sheet', 'HeatMap', 'Heat Map'" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    diagnostics = diagnostics & "✓ Found HeatMap Sheet: " & wsHeatMap.Name & vbCrLf & vbCrLf
    
    ' Step 3: Find Status column in HeatMap Sheet
    diagnostics = diagnostics & "=== STEP 3: Finding Status Column in HeatMap ===" & vbCrLf
    statusCol = FindStatusColumn(wsHeatMap)
    
    If statusCol = 0 Then
        MsgBox "ERROR: 'Status' column not found in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Please ensure there's a column header containing 'Status' or 'Current Status'." & vbCrLf & vbCrLf & _
               "Found headers:" & vbCrLf & ListHeatMapHeaders(wsHeatMap), vbCritical, "Column Not Found"
        Exit Sub
    End If
    diagnostics = diagnostics & "✓ Found Status column: " & ColumnLetter(statusCol) & " (" & statusCol & ")" & vbCrLf
    diagnostics = diagnostics & "  Header: " & wsHeatMap.Cells(1, statusCol).Value & vbCrLf & vbCrLf
    
    ' Step 4: Scan Evaluation Results for operation codes and statuses
    diagnostics = diagnostics & "=== STEP 4: Scanning Evaluation Results ===" & vbCrLf
    evalLastRow = wsEval.Cells(wsEval.Rows.Count, 1).End(xlUp).Row
    diagnostics = diagnostics & "Last row in Evaluation Results: " & evalLastRow & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    Dim overallStatusRow As Long
    Dim summaryRow As Long
    overallStatusRow = FindSectionRow(wsEval, "Overall Status by Op Code")
    summaryRow = FindSectionRow(wsEval, "Operation Mode Summary")
    
    If overallStatusRow = 0 And summaryRow = 0 Then
        MsgBox "ERROR: Could not find evaluation data sections!" & vbCrLf & vbCrLf & _
               "Looking for:" & vbCrLf & _
               "  • 'Overall Status by Op Code'" & vbCrLf & _
               "  • 'Operation Mode Summary'" & vbCrLf & vbCrLf & _
               "First 20 rows:" & vbCrLf & ListEvaluationRows(wsEval, 20), vbCritical, "Data Not Found"
        Exit Sub
    End If
    
    diagnostics = diagnostics & "✓ Found 'Overall Status by Op Code' at row: " & overallStatusRow & vbCrLf
    diagnostics = diagnostics & "✓ Found 'Operation Mode Summary' at row: " & summaryRow & vbCrLf & vbCrLf
    
    ' Step 5: Get HeatMap operation codes
    diagnostics = diagnostics & "=== STEP 5: Reading HeatMap Operation Codes ===" & vbCrLf
    heatMapLastRow = wsHeatMap.Cells(wsHeatMap.Rows.Count, 1).End(xlUp).Row
    diagnostics = diagnostics & "Last row in HeatMap: " & heatMapLastRow & vbCrLf
    diagnostics = diagnostics & "First 10 HeatMap codes:" & vbCrLf & ListHeatMapCodes(wsHeatMap, 10) & vbCrLf
    
    ' Step 6: Process each operation in HeatMap
    diagnostics = diagnostics & "=== STEP 6: Updating HeatMap Statuses ===" & vbCrLf
    
    For i = 2 To heatMapLastRow ' Start from row 2 (skip header)
        opCode = Trim(wsHeatMap.Cells(i, 1).Value)
        
        If opCode <> "" And IsNumeric(opCode) Then
            ' Look for this code in Evaluation Results
            status = FindStatusForCode(wsEval, opCode, overallStatusRow, summaryRow, evalLastRow)
            
            If status <> "" Then
                ' Update the status cell with colored dot
                UpdateStatusCell wsHeatMap.Cells(i, statusCol), status
                updated = updated + 1
                
                ' Log first 5 updates
                If updated <= 5 Then
                    diagnostics = diagnostics & "  ✓ Updated row " & i & ": " & opCode & " → " & status & vbCrLf
                End If
            Else
                notFound = notFound + 1
                
                ' Log first 3 not found
                If notFound <= 3 Then
                    diagnostics = diagnostics & "  ✗ Not found row " & i & ": " & opCode & vbCrLf
                End If
            End If
        End If
    Next i
    
    ' Step 7: Show results
    Dim duration As Double
    duration = Round(Timer - startTime, 2)
    
    diagnostics = diagnostics & vbCrLf & "=== STEP 7: Summary ===" & vbCrLf
    diagnostics = diagnostics & "Operations updated: " & updated & vbCrLf
    diagnostics = diagnostics & "Operations not found: " & notFound & vbCrLf
    diagnostics = diagnostics & "Time taken: " & duration & " seconds" & vbCrLf
    
    ' Show diagnostic information
    MsgBox diagnostics, vbInformation, "HeatMap Update Complete"
    
    ' Also show summary
    MsgBox "HeatMap Status Update Complete!" & vbCrLf & vbCrLf & _
           "✓ Updated: " & updated & " operations" & vbCrLf & _
           "✗ Not found: " & notFound & " operations" & vbCrLf & vbCrLf & _
           "Time: " & duration & " seconds", vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Diagnostics so far:" & vbCrLf & diagnostics, vbCritical, "Error Occurred"
End Sub

' Find section row by looking for section header
Function FindSectionRow(ws As Worksheet, sectionName As String) As Long
    Dim i As Long
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
        If InStr(1, ws.Cells(i, 1).Value, sectionName, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
    
    FindSectionRow = 0
End Function

' Find status for a specific operation code
Function FindStatusForCode(ws As Worksheet, opCode As String, overallStartRow As Long, summaryStartRow As Long, lastRow As Long) As String
    Dim i As Long
    Dim cellCode As String
    Dim statusColIndex As Long
    
    ' Search in Overall Status section first (rows after overallStartRow)
    If overallStartRow > 0 Then
        For i = overallStartRow + 1 To lastRow
            ' Stop if we hit another section
            If ws.Cells(i, 1).Value = "Operation Mode Summary" Then Exit For
            
            cellCode = Trim(ws.Cells(i, 1).Value)
            If cellCode = opCode Then
                ' Look for status in adjacent columns (typically column C for Overall Status)
                statusColIndex = FindStatusColumnInRow(ws, i)
                If statusColIndex > 0 Then
                    FindStatusForCode = Trim(ws.Cells(i, statusColIndex).Value)
                    Exit Function
                End If
            End If
        Next i
    End If
    
    ' Search in Operation Mode Summary section
    If summaryStartRow > 0 Then
        For i = summaryStartRow + 1 To lastRow
            cellCode = Trim(ws.Cells(i, 1).Value)  ' Op Code typically in column A or F
            
            ' Check column A
            If cellCode = opCode Then
                statusColIndex = FindStatusColumnInRow(ws, i)
                If statusColIndex > 0 Then
                    FindStatusForCode = Trim(ws.Cells(i, statusColIndex).Value)
                    Exit Function
                End If
            End If
            
            ' Also check column F (where Op Code might be in summary section)
            If Trim(ws.Cells(i, 6).Value) = opCode Then
                ' Status typically in column I for summary
                If ws.Cells(i, 9).Value <> "" Then
                    FindStatusForCode = Trim(ws.Cells(i, 9).Value)
                    Exit Function
                End If
            End If
        Next i
    End If
    
    FindStatusForCode = ""
End Function

' Find status column in a specific row (look for column with status values)
Function FindStatusColumnInRow(ws As Worksheet, row As Long) As Long
    Dim col As Long
    Dim cellValue As String
    
    ' Check columns B through M for status values
    For col = 2 To 13
        cellValue = UCase(Trim(ws.Cells(row, col).Value))
        If cellValue = "RED" Or cellValue = "YELLOW" Or cellValue = "GREEN" Or cellValue = "N/A" Then
            FindStatusColumnInRow = col
            Exit Function
        End If
    Next col
    
    FindStatusColumnInRow = 0
End Function

' Find Status column in HeatMap sheet
Function FindStatusColumn(ws As Worksheet) As Long
    Dim col As Long
    Dim headerValue As String
    
    ' Check first row for Status header
    For col = 1 To 50
        headerValue = UCase(Trim(ws.Cells(1, col).Value))
        If InStr(1, headerValue, "STATUS", vbTextCompare) > 0 Then
            FindStatusColumn = col
            Exit Function
        End If
    Next col
    
    FindStatusColumn = 0
End Function

' Update status cell with colored dot
Sub UpdateStatusCell(cell As Range, status As String)
    Dim statusUpper As String
    statusUpper = UCase(Trim(status))
    
    ' Set the cell value to a filled circle using Wingdings font
    cell.Value = "l"  ' This is a filled circle in Wingdings
    cell.Font.Name = "Wingdings"
    cell.Font.Size = 14
    cell.HorizontalAlignment = xlCenter
    cell.VerticalAlignment = xlCenter
    
    ' Set color based on status
    Select Case statusUpper
        Case "RED"
            cell.Font.Color = RGB(255, 0, 0)  ' Red
        Case "YELLOW"
            cell.Font.Color = RGB(255, 192, 0)  ' Yellow/Orange
        Case "GREEN"
            cell.Font.Color = RGB(0, 176, 80)  ' Green
        Case "N/A", ""
            cell.Font.Color = RGB(128, 128, 128)  ' Gray
        Case Else
            cell.Font.Color = RGB(0, 0, 0)  ' Black for unknown
    End Select
End Sub

' Helper: List all sheets in workbook
Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim result As String
    
    For Each ws In ThisWorkbook.Worksheets
        result = result & "  • " & ws.Name & vbCrLf
    Next ws
    
    ListAllSheets = result
End Function

' Helper: List HeatMap headers
Function ListHeatMapHeaders(ws As Worksheet) As String
    Dim col As Long
    Dim result As String
    
    For col = 1 To 20
        If ws.Cells(1, col).Value <> "" Then
            result = result & "  Column " & ColumnLetter(col) & ": " & ws.Cells(1, col).Value & vbCrLf
        End If
    Next col
    
    ListHeatMapHeaders = result
End Function

' Helper: List first N evaluation rows
Function ListEvaluationRows(ws As Worksheet, numRows As Long) As String
    Dim i As Long
    Dim result As String
    
    For i = 1 To numRows
        If ws.Cells(i, 1).Value <> "" Then
            result = result & "  Row " & i & ": " & ws.Cells(i, 1).Value & vbCrLf
        End If
    Next i
    
    ListEvaluationRows = result
End Function

' Helper: List first N HeatMap codes
Function ListHeatMapCodes(ws As Worksheet, numRows As Long) As String
    Dim i As Long
    Dim result As String
    
    For i = 2 To numRows + 1
        If ws.Cells(i, 1).Value <> "" Then
            result = result & "  Row " & i & ": " & ws.Cells(i, 1).Value & vbCrLf
        End If
    Next i
    
    ListHeatMapCodes = result
End Function

' Helper: Get column letter from number
Function ColumnLetter(colNum As Long) As String
    Dim dividend As Long
    Dim modulo As Long
    Dim result As String
    
    dividend = colNum
    
    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        result = Chr(65 + modulo) & result
        dividend = (dividend - modulo) \ 26
    Loop
    
    ColumnLetter = result
End Function

' Optional: Create a button on HeatMap Sheet
Sub CreateUpdateButtonOnHeatMap()
    Dim ws As Worksheet
    Dim btn As Button
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("HeatMap")
    If ws Is Nothing Then Set ws = ThisWorkbook.Sheets("Heat Map")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if it exists
    On Error Resume Next
    ws.Buttons("Update HeatMap Status").Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = "Update HeatMap Status"
    btn.Caption = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatusEnhanced"
    
    MsgBox "Button created on " & ws.Name & "!" & vbCrLf & vbCrLf & _
           "Click the button to update HeatMap statuses.", vbInformation, "Button Created"
End Sub
