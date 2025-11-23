Attribute VB_Name = "HeatMapUpdate_Module_COMPREHENSIVE_DIAGNOSTIC"
'===============================================================================
' Module: HeatMapUpdate_Module_COMPREHENSIVE_DIAGNOSTIC
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive diagnostics
' Version: 4.0 - Enhanced diagnostic version
' Author: GitHub Copilot
' Date: 2025-11-23
'===============================================================================

Option Explicit

' Module-level constants for configuration
Private Const EVAL_SHEET_NAME As String = "Evaluation Results"
Private Const HEATMAP_SHEET_NAME As String = "HeatMap Sheet"
Private Const OP_CODE_COL As Long = 1 ' Column A
Private Const STATUS_COL As Long = 3  ' Column C - adjust if different

'===============================================================================
' Main Function: UpdateHeatMapStatus
' Purpose: Transfer evaluation results to HeatMap Sheet with detailed diagnostics
'===============================================================================
Public Sub UpdateHeatMapStatus()
    On Error GoTo ErrorHandler
    
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim diagMsg As String
    Dim updatedCount As Long
    Dim startTime As Double
    
    startTime = Timer
    diagMsg = "=== HEATMAP UPDATE DIAGNOSTIC REPORT ===" & vbCrLf & vbCrLf
    
    ' Step 1: Verify Evaluation Results sheet exists
    diagMsg = diagMsg & "STEP 1: Checking Evaluation Results sheet..." & vbCrLf
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets(EVAL_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        diagMsg = diagMsg & "❌ ERROR: Sheet '" & EVAL_SHEET_NAME & "' not found!" & vbCrLf
        diagMsg = diagMsg & vbCrLf & "Available sheets:" & vbCrLf
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Sheets
            diagMsg = diagMsg & "  - " & ws.Name & vbCrLf
        Next ws
        MsgBox diagMsg, vbExclamation, "Sheet Not Found"
        Exit Sub
    End If
    diagMsg = diagMsg & "✓ Found: " & EVAL_SHEET_NAME & vbCrLf
    
    ' Step 2: Verify HeatMap Sheet exists
    diagMsg = diagMsg & vbCrLf & "STEP 2: Checking HeatMap Sheet..." & vbCrLf
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets(HEATMAP_SHEET_NAME)
    On Error GoTo ErrorHandler
    
    If wsHeatMap Is Nothing Then
        diagMsg = diagMsg & "❌ ERROR: Sheet '" & HEATMAP_SHEET_NAME & "' not found!" & vbCrLf
        MsgBox diagMsg, vbExclamation, "Sheet Not Found"
        Exit Sub
    End If
    diagMsg = diagMsg & "✓ Found: " & HEATMAP_SHEET_NAME & vbCrLf
    
    ' Step 3: Analyze Evaluation Results structure
    diagMsg = diagMsg & vbCrLf & "STEP 3: Analyzing Evaluation Results structure..." & vbCrLf
    Dim evalRowCount As Long
    Dim overallStatusRow As Long
    Dim summaryRow As Long
    
    evalRowCount = wsEval.Cells(wsEval.Rows.Count, OP_CODE_COL).End(xlUp).Row
    diagMsg = diagMsg & "Last row with data: " & evalRowCount & vbCrLf
    
    ' Find "Overall Status by Op Code" row
    overallStatusRow = FindRowWithText(wsEval, "Overall Status by Op Code")
    If overallStatusRow > 0 Then
        diagMsg = diagMsg & "✓ Found 'Overall Status by Op Code' at row: " & overallStatusRow & vbCrLf
    Else
        diagMsg = diagMsg & "⚠ 'Overall Status by Op Code' section not found" & vbCrLf
    End If
    
    ' Find "Operation Mode Summary" row
    summaryRow = FindRowWithText(wsEval, "Operation Mode Summary")
    If summaryRow > 0 Then
        diagMsg = diagMsg & "✓ Found 'Operation Mode Summary' at row: " & summaryRow & vbCrLf
    Else
        diagMsg = diagMsg & "⚠ 'Operation Mode Summary' section not found" & vbCrLf
    End If
    
    ' Step 4: Analyze HeatMap structure
    diagMsg = diagMsg & vbCrLf & "STEP 4: Analyzing HeatMap Sheet structure..." & vbCrLf
    Dim heatMapRowCount As Long
    Dim statusColIndex As Long
    
    heatMapRowCount = wsHeatMap.Cells(wsHeatMap.Rows.Count, OP_CODE_COL).End(xlUp).Row
    diagMsg = diagMsg & "HeatMap rows with data: " & heatMapRowCount & vbCrLf
    
    ' Find Status column
    statusColIndex = FindStatusColumn(wsHeatMap)
    If statusColIndex > 0 Then
        diagMsg = diagMsg & "✓ Found 'Status' column at: " & ColumnLetter(statusColIndex) & vbCrLf
    Else
        diagMsg = diagMsg & "❌ 'Status' column not found! Using column C as default." & vbCrLf
        statusColIndex = STATUS_COL
    End If
    
    ' Step 5: Sample data from both sheets
    diagMsg = diagMsg & vbCrLf & "STEP 5: Sample data from sheets..." & vbCrLf
    diagMsg = diagMsg & vbCrLf & "Sample from Evaluation Results (rows after 'Overall Status'):" & vbCrLf
    If overallStatusRow > 0 Then
        Dim sampleRow As Long
        For sampleRow = overallStatusRow + 1 To Application.Min(overallStatusRow + 5, evalRowCount)
            Dim opCode As String
            Dim opName As String
            Dim status As String
            opCode = Trim(wsEval.Cells(sampleRow, 1).Value)
            opName = Trim(wsEval.Cells(sampleRow, 2).Value)
            status = Trim(wsEval.Cells(sampleRow, 3).Value)
            If opCode <> "" Then
                diagMsg = diagMsg & "  Row " & sampleRow & ": [" & opCode & "] " & opName & " → " & status & vbCrLf
            End If
        Next sampleRow
    End If
    
    diagMsg = diagMsg & vbCrLf & "Sample from HeatMap Sheet:" & vbCrLf
    For sampleRow = 2 To Application.Min(7, heatMapRowCount)
        opCode = Trim(wsHeatMap.Cells(sampleRow, OP_CODE_COL).Value)
        If opCode <> "" And IsNumeric(opCode) Then
            diagMsg = diagMsg & "  Row " & sampleRow & ": [" & opCode & "]" & vbCrLf
        End If
    Next sampleRow
    
    ' Step 6: Perform the update
    diagMsg = diagMsg & vbCrLf & "STEP 6: Updating statuses..." & vbCrLf
    updatedCount = TransferStatuses(wsEval, wsHeatMap, overallStatusRow, summaryRow, statusColIndex, diagMsg)
    
    ' Step 7: Summary
    Dim elapsed As Double
    elapsed = Timer - startTime
    diagMsg = diagMsg & vbCrLf & "=== UPDATE COMPLETE ===" & vbCrLf
    diagMsg = diagMsg & "Operations updated: " & updatedCount & vbCrLf
    diagMsg = diagMsg & "Time elapsed: " & Format(elapsed, "0.00") & " seconds" & vbCrLf
    
    If updatedCount = 0 Then
        diagMsg = diagMsg & vbCrLf & "⚠ WARNING: No operations were updated!" & vbCrLf
        diagMsg = diagMsg & "Possible reasons:" & vbCrLf
        diagMsg = diagMsg & "  1. Operation codes in Evaluation Results don't match HeatMap Sheet" & vbCrLf
        diagMsg = diagMsg & "  2. Evaluation hasn't been run yet (no data in Evaluation Results)" & vbCrLf
        diagMsg = diagMsg & "  3. Status column location is incorrect" & vbCrLf
    End If
    
    ' Display diagnostic report
    MsgBox diagMsg, vbInformation, "HeatMap Update Diagnostic Report"
    
    Exit Sub

ErrorHandler:
    diagMsg = diagMsg & vbCrLf & "❌ FATAL ERROR: " & Err.Description & vbCrLf
    diagMsg = diagMsg & "Error Number: " & Err.Number & vbCrLf
    MsgBox diagMsg, vbCritical, "Error in HeatMap Update"
End Sub

'===============================================================================
' Function: TransferStatuses
' Purpose: Transfer status values from Evaluation to HeatMap with matching
'===============================================================================
Private Function TransferStatuses(wsEval As Worksheet, wsHeatMap As Worksheet, _
                                 overallStatusRow As Long, summaryRow As Long, _
                                 statusColIndex As Long, ByRef diagMsg As String) As Long
    On Error GoTo ErrorHandler
    
    Dim evalRow As Long
    Dim heatMapRow As Long
    Dim opCode As String
    Dim status As String
    Dim matchedRow As Long
    Dim updateCount As Long
    Dim evalLastRow As Long
    Dim heatMapLastRow As Long
    
    updateCount = 0
    evalLastRow = wsEval.Cells(wsEval.Rows.Count, 1).End(xlUp).Row
    heatMapLastRow = wsHeatMap.Cells(wsHeatMap.Rows.Count, OP_CODE_COL).End(xlUp).Row
    
    diagMsg = diagMsg & "Scanning Evaluation Results rows " & (overallStatusRow + 1) & " to " & evalLastRow & vbCrLf
    
    ' Process rows after "Overall Status by Op Code"
    If overallStatusRow > 0 Then
        For evalRow = overallStatusRow + 1 To evalLastRow
            opCode = Trim(wsEval.Cells(evalRow, 1).Value)
            
            ' Skip if not a valid operation code or if it's a section header
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                status = Trim(wsEval.Cells(evalRow, 3).Value)
                
                ' Find matching row in HeatMap
                matchedRow = FindOperationInHeatMap(wsHeatMap, opCode, heatMapLastRow)
                
                If matchedRow > 0 And status <> "" Then
                    ' Update the status
                    UpdateStatusCell wsHeatMap, matchedRow, statusColIndex, status
                    updateCount = updateCount + 1
                    
                    If updateCount <= 5 Then ' Show first 5 matches
                        diagMsg = diagMsg & "  ✓ Matched [" & opCode & "] → Row " & matchedRow & " = " & status & vbCrLf
                    End If
                End If
            End If
            
            ' Stop if we reach the Operation Mode Summary section
            If summaryRow > 0 And evalRow >= summaryRow - 1 Then Exit For
        Next evalRow
    End If
    
    ' Process Operation Mode Summary section if it exists
    If summaryRow > 0 Then
        diagMsg = diagMsg & vbCrLf & "Processing Operation Mode Summary section..." & vbCrLf
        
        For evalRow = summaryRow + 1 To evalLastRow
            opCode = Trim(wsEval.Cells(evalRow, 1).Value)
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                ' Look for status in different columns - Operation Mode Summary has different structure
                ' Try column I (9) for Final Status in summary
                status = Trim(wsEval.Cells(evalRow, 9).Value)
                If status = "" Then
                    status = Trim(wsEval.Cells(evalRow, 3).Value)
                End If
                
                matchedRow = FindOperationInHeatMap(wsHeatMap, opCode, heatMapLastRow)
                
                If matchedRow > 0 And status <> "" Then
                    UpdateStatusCell wsHeatMap, matchedRow, statusColIndex, status
                    updateCount = updateCount + 1
                    
                    If updateCount <= 10 Then ' Show additional matches from summary
                        diagMsg = diagMsg & "  ✓ Summary [" & opCode & "] → Row " & matchedRow & " = " & status & vbCrLf
                    End If
                End If
            End If
        Next evalRow
    End If
    
    TransferStatuses = updateCount
    Exit Function

ErrorHandler:
    diagMsg = diagMsg & "Error in TransferStatuses: " & Err.Description & vbCrLf
    TransferStatuses = updateCount
End Function

'===============================================================================
' Function: FindOperationInHeatMap
' Purpose: Find row with matching operation code in HeatMap sheet
'===============================================================================
Private Function FindOperationInHeatMap(ws As Worksheet, opCode As String, _
                                       lastRow As Long) As Long
    Dim i As Long
    Dim cellValue As String
    
    For i = 2 To lastRow ' Start from row 2 (skip header)
        cellValue = Trim(ws.Cells(i, OP_CODE_COL).Value)
        If cellValue = opCode Then
            FindOperationInHeatMap = i
            Exit Function
        End If
    Next i
    
    FindOperationInHeatMap = 0 ' Not found
End Function

'===============================================================================
' Function: UpdateStatusCell
' Purpose: Update status cell with colored dot based on status
'===============================================================================
Private Sub UpdateStatusCell(ws As Worksheet, rowNum As Long, colNum As Long, status As String)
    Dim cell As Range
    Set cell = ws.Cells(rowNum, colNum)
    
    status = UCase(Trim(status))
    
    ' Set the text to a filled circle (using Wingdings font)
    cell.Value = "●"
    cell.Font.Name = "Wingdings"
    cell.Font.Size = 14
    cell.HorizontalAlignment = xlCenter
    
    ' Set color based on status
    Select Case status
        Case "RED"
            cell.Font.Color = RGB(255, 0, 0) ' Red
        Case "YELLOW"
            cell.Font.Color = RGB(255, 192, 0) ' Yellow/Orange
        Case "GREEN"
            cell.Font.Color = RGB(0, 176, 80) ' Green
        Case Else
            cell.Font.Color = RGB(128, 128, 128) ' Gray for N/A
    End Select
End Sub

'===============================================================================
' Function: FindRowWithText
' Purpose: Find row containing specific text in column A
'===============================================================================
Private Function FindRowWithText(ws As Worksheet, searchText As String) As Long
    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To lastRow
        cellValue = Trim(ws.Cells(i, 1).Value)
        If InStr(1, cellValue, searchText, vbTextCompare) > 0 Then
            FindRowWithText = i
            Exit Function
        End If
    Next i
    
    FindRowWithText = 0 ' Not found
End Function

'===============================================================================
' Function: FindStatusColumn
' Purpose: Find column with "Status" header in row 1
'===============================================================================
Private Function FindStatusColumn(ws As Worksheet) As Long
    Dim lastCol As Long
    Dim i As Long
    Dim cellValue As String
    
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    For i = 1 To lastCol
        cellValue = Trim(ws.Cells(1, i).Value)
        If InStr(1, cellValue, "Status", vbTextCompare) > 0 Then
            FindStatusColumn = i
            Exit Function
        End If
    Next i
    
    FindStatusColumn = 0 ' Not found
End Function

'===============================================================================
' Function: ColumnLetter
' Purpose: Convert column number to letter (e.g., 1 → A, 27 → AA)
'===============================================================================
Private Function ColumnLetter(colNum As Long) As String
    Dim col As Long
    Dim letter As String
    
    col = colNum
    Do While col > 0
        letter = Chr(((col - 1) Mod 26) + 65) & letter
        col = (col - 1) \ 26
    Loop
    
    ColumnLetter = letter
End Function

'===============================================================================
' Sub: CreateUpdateButton
' Purpose: Create a button on HeatMap Sheet to run the update
'===============================================================================
Public Sub CreateUpdateButton()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim btn As Button
    
    ' Get HeatMap Sheet
    Set ws = ThisWorkbook.Sheets(HEATMAP_SHEET_NAME)
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing button if it exists
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "btnUpdateHeatMap" Then
            shp.Delete
        End If
    Next shp
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = "btnUpdateHeatMap"
    btn.Text = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on " & HEATMAP_SHEET_NAME & "!", vbInformation
End Sub
