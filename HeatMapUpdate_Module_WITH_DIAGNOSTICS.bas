Attribute VB_Name = "HeatMapUpdate_Diagnostic"
' ====================================================================
' Module: HeatMapUpdate_Diagnostic
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive diagnostics
' Features: Detailed error messages, step-by-step diagnostics, data validation
' ====================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results with full diagnostics
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
    Dim summaryStartRow As Long
    Dim evalOpCodes As String
    Dim heatMapOpCodes As String
    Dim matchedCodes As String
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugInfo = "=== HEATMAP UPDATE DIAGNOSTIC REPORT ===" & vbCrLf & vbCrLf
    
    ' Step 1: Verify sheets exist
    debugInfo = debugInfo & "STEP 1: Checking for required sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetSheetList(), _
               vbCritical, "Missing Sheet"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ 'Evaluation Results' sheet found" & vbCrLf
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetSheetList(), _
               vbCritical, "Missing Sheet"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ 'HeatMap Sheet' found" & vbCrLf & vbCrLf
    
    ' Step 2: Find data sections in Evaluation Results
    debugInfo = debugInfo & "STEP 2: Analyzing Evaluation Results structure..." & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    Dim overallStatusRow As Long
    overallStatusRow = FindRowWithText(wsEval, "Overall Status by Op Code", 1, 100)
    
    If overallStatusRow = 0 Then
        debugInfo = debugInfo & "  ✗ 'Overall Status by Op Code' section NOT FOUND" & vbCrLf
    Else
        debugInfo = debugInfo & "  ✓ 'Overall Status by Op Code' found at row " & overallStatusRow & vbCrLf
    End If
    
    ' Find "Operation Mode Summary" section
    summaryStartRow = FindRowWithText(wsEval, "Operation Mode Summary", 1, 200)
    
    If summaryStartRow = 0 Then
        debugInfo = debugInfo & "  ✗ 'Operation Mode Summary' section NOT FOUND" & vbCrLf
    Else
        debugInfo = debugInfo & "  ✓ 'Operation Mode Summary' found at row " & summaryStartRow & vbCrLf
    End If
    
    ' Get last row with data in Evaluation Results
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  • Last row in Evaluation Results: " & lastRowEval & vbCrLf & vbCrLf
    
    ' Step 3: Analyze HeatMap Sheet structure
    debugInfo = debugInfo & "STEP 3: Analyzing HeatMap Sheet structure..." & vbCrLf
    
    ' Find "Op Code" column in HeatMap
    Dim heatMapCodeCol As Long
    heatMapCodeCol = FindColumnWithText(wsHeatMap, "Op Code", 1)
    
    If heatMapCodeCol = 0 Then
        heatMapCodeCol = 1 ' Default to column A
        debugInfo = debugInfo & "  • Using column A for Op Codes (default)" & vbCrLf
    Else
        debugInfo = debugInfo & "  ✓ 'Op Code' column found at column " & ColumnLetter(heatMapCodeCol) & vbCrLf
    End If
    
    ' Find "Status" or "Current Status" column
    Dim heatMapStatusCol As Long
    heatMapStatusCol = FindColumnWithText(wsHeatMap, "Status", 1)
    If heatMapStatusCol = 0 Then
        heatMapStatusCol = FindColumnWithText(wsHeatMap, "Current Status", 1)
    End If
    
    If heatMapStatusCol = 0 Then
        debugInfo = debugInfo & "  ✗ 'Status' column NOT FOUND in HeatMap Sheet" & vbCrLf
        MsgBox debugInfo, vbExclamation, "Diagnostic Information"
        Exit Sub
    Else
        debugInfo = debugInfo & "  ✓ 'Status' column found at column " & ColumnLetter(heatMapStatusCol) & vbCrLf
    End If
    
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, heatMapCodeCol).End(xlUp).Row
    debugInfo = debugInfo & "  • Last row in HeatMap Sheet: " & lastRowHeatMap & vbCrLf & vbCrLf
    
    ' Step 4: Sample operation codes
    debugInfo = debugInfo & "STEP 4: Sampling operation codes..." & vbCrLf
    
    debugInfo = debugInfo & "  Evaluation Results Op Codes (first 5):" & vbCrLf
    For i = 2 To Application.Min(6, lastRowEval)
        Dim evalCode As String
        evalCode = Trim(wsEval.Cells(i, 1).Value)
        If evalCode <> "" Then
            evalOpCodes = evalOpCodes & evalCode & ", "
            debugInfo = debugInfo & "    - Row " & i & ": " & evalCode & vbCrLf
        End If
    Next i
    
    debugInfo = debugInfo & vbCrLf & "  HeatMap Sheet Op Codes (first 5):" & vbCrLf
    For i = 2 To Application.Min(6, lastRowHeatMap)
        Dim heatCode As String
        heatCode = Trim(wsHeatMap.Cells(i, heatMapCodeCol).Value)
        If heatCode <> "" Then
            heatMapOpCodes = heatMapOpCodes & heatCode & ", "
            debugInfo = debugInfo & "    - Row " & i & ": " & heatCode & vbCrLf
        End If
    Next i
    debugInfo = debugInfo & vbCrLf
    
    ' Step 5: Attempt to match and update
    debugInfo = debugInfo & "STEP 5: Attempting to match and update statuses..." & vbCrLf & vbCrLf
    
    Dim matchCount As Long
    matchCount = 0
    
    ' Process sub-operations from "Overall Status by Op Code" section
    If overallStatusRow > 0 Then
        Dim overallEndRow As Long
        overallEndRow = overallStatusRow + 1
        
        ' Find the end of this section
        Do While overallEndRow < lastRowEval
            If Trim(wsEval.Cells(overallEndRow, 1).Value) = "" And _
               Trim(wsEval.Cells(overallEndRow + 1, 1).Value) = "" Then
                Exit Do
            End If
            overallEndRow = overallEndRow + 1
        Loop
        
        debugInfo = debugInfo & "Processing Overall Status section (rows " & (overallStatusRow + 1) & " to " & overallEndRow & ")..." & vbCrLf
        
        For i = overallStatusRow + 1 To overallEndRow
            opCode = Trim(wsEval.Cells(i, 1).Value)
            
            If opCode <> "" And IsNumeric(opCode) Then
                ' Find "Overall Status" column (column C typically)
                finalStatus = Trim(wsEval.Cells(i, 3).Value)
                
                If finalStatus <> "" And finalStatus <> "Overall Status" Then
                    ' Try to find this operation in HeatMap
                    For j = 2 To lastRowHeatMap
                        If Trim(wsHeatMap.Cells(j, heatMapCodeCol).Value) = opCode Then
                            ' Update the status
                            wsHeatMap.Cells(j, heatMapStatusCol).Value = GetStatusDot(finalStatus)
                            FormatStatusCell wsHeatMap.Cells(j, heatMapStatusCol), finalStatus
                            
                            updatedCount = updatedCount + 1
                            matchCount = matchCount + 1
                            matchedCodes = matchedCodes & opCode & ", "
                            
                            If matchCount <= 3 Then
                                debugInfo = debugInfo & "  ✓ Matched: " & opCode & " → " & finalStatus & vbCrLf
                            End If
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i
        
        debugInfo = debugInfo & "  • Processed " & matchCount & " operations from Overall Status section" & vbCrLf & vbCrLf
    End If
    
    ' Process parent operations from "Operation Mode Summary" section
    If summaryStartRow > 0 Then
        Dim summaryEndRow As Long
        summaryEndRow = summaryStartRow + 1
        
        ' Find the end of this section
        Do While summaryEndRow < lastRowEval
            If Trim(wsEval.Cells(summaryEndRow, 1).Value) = "" And _
               Trim(wsEval.Cells(summaryEndRow + 1, 1).Value) = "" Then
                Exit Do
            End If
            summaryEndRow = summaryEndRow + 1
        Loop
        
        debugInfo = debugInfo & "Processing Operation Mode Summary section (rows " & (summaryStartRow + 1) & " to " & summaryEndRow & ")..." & vbCrLf
        
        Dim summaryMatchCount As Long
        summaryMatchCount = 0
        
        For i = summaryStartRow + 1 To summaryEndRow
            ' Try column F first (Op Code in summary section)
            opCode = Trim(wsEval.Cells(i, 6).Value)
            
            If opCode <> "" And IsNumeric(opCode) Then
                ' Find "Final Status" column (column I typically in summary)
                finalStatus = Trim(wsEval.Cells(i, 9).Value)
                
                If finalStatus <> "" And finalStatus <> "Final Status" Then
                    ' Try to find this operation in HeatMap
                    For j = 2 To lastRowHeatMap
                        If Trim(wsHeatMap.Cells(j, heatMapCodeCol).Value) = opCode Then
                            ' Update the status
                            wsHeatMap.Cells(j, heatMapStatusCol).Value = GetStatusDot(finalStatus)
                            FormatStatusCell wsHeatMap.Cells(j, heatMapStatusCol), finalStatus
                            
                            updatedCount = updatedCount + 1
                            summaryMatchCount = summaryMatchCount + 1
                            matchedCodes = matchedCodes & opCode & ", "
                            
                            If summaryMatchCount <= 3 Then
                                debugInfo = debugInfo & "  ✓ Matched: " & opCode & " → " & finalStatus & vbCrLf
                            End If
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i
        
        debugInfo = debugInfo & "  • Processed " & summaryMatchCount & " operations from Summary section" & vbCrLf & vbCrLf
    End If
    
    ' Final summary
    debugInfo = debugInfo & "=== FINAL RESULTS ===" & vbCrLf
    debugInfo = debugInfo & "Total operations updated: " & updatedCount & vbCrLf
    debugInfo = debugInfo & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf & vbCrLf
    
    If updatedCount = 0 Then
        debugInfo = debugInfo & "⚠ WARNING: No operations were updated!" & vbCrLf & vbCrLf
        debugInfo = debugInfo & "POSSIBLE ISSUES:" & vbCrLf
        debugInfo = debugInfo & "1. Operation codes don't match between sheets" & vbCrLf
        debugInfo = debugInfo & "2. Status columns are in different positions than expected" & vbCrLf
        debugInfo = debugInfo & "3. Data sections not found in expected locations" & vbCrLf
        debugInfo = debugInfo & "4. Sheet structure is different than expected" & vbCrLf
        
        MsgBox debugInfo, vbExclamation, "No Updates Made - Diagnostic Report"
    Else
        debugInfo = debugInfo & "✓ Successfully updated " & updatedCount & " operation statuses!"
        MsgBox debugInfo, vbInformation, "Update Complete - Diagnostic Report"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error occurred: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl, vbCritical, "Error"
End Sub

' Helper function to find row containing specific text
Private Function FindRowWithText(ws As Worksheet, searchText As String, startRow As Long, endRow As Long) As Long
    Dim i As Long
    FindRowWithText = 0
    
    For i = startRow To endRow
        If InStr(1, ws.Cells(i, 1).Value, searchText, vbTextCompare) > 0 Then
            FindRowWithText = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find column containing specific text in row 1
Private Function FindColumnWithText(ws As Worksheet, searchText As String, searchRow As Long) As Long
    Dim i As Long
    FindColumnWithText = 0
    
    For i = 1 To 50 ' Search first 50 columns
        If InStr(1, ws.Cells(searchRow, i).Value, searchText, vbTextCompare) > 0 Then
            FindColumnWithText = i
            Exit Function
        End If
    Next i
End Function

' Helper function to get column letter from number
Private Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

' Helper function to get list of all sheets
Private Function GetSheetList() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    For Each ws In ThisWorkbook.Worksheets
        sheetList = sheetList & "  • " & ws.Name & vbCrLf
    Next ws
    
    GetSheetList = sheetList
End Function

' Get colored dot character based on status
Private Function GetStatusDot(status As String) As String
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusDot = "●"
        Case "YELLOW"
            GetStatusDot = "●"
        Case "GREEN"
            GetStatusDot = "●"
        Case "N/A", ""
            GetStatusDot = "●"
        Case Else
            GetStatusDot = "●"
    End Select
End Function

' Format cell with appropriate color
Private Sub FormatStatusCell(cell As Range, status As String)
    With cell
        .Font.Name = "Wingdings"
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
        
        Select Case UCase(Trim(status))
            Case "RED"
                .Font.Color = RGB(255, 0, 0)     ' Red
            Case "YELLOW"
                .Font.Color = RGB(255, 192, 0)   ' Yellow/Orange
            Case "GREEN"
                .Font.Color = RGB(0, 176, 80)    ' Green
            Case "N/A", ""
                .Font.Color = RGB(128, 128, 128) ' Gray
            Case Else
                .Font.Color = RGB(0, 0, 0)       ' Black
        End Select
    End With
End Sub

' Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Remove existing button if present
    btnName = "UpdateHeatMapButton"
    On Error Resume Next
    ws.Buttons(btnName).Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    With btn
        .Name = btnName
        .Text = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update statuses after running evaluation.", _
           vbInformation, "Button Created"
End Sub
