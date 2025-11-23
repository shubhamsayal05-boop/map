Attribute VB_Name = "HeatMapUpdate_Enhanced"
' ====================================================================
' Module: HeatMapUpdate_Enhanced
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Version: Enhanced Debug (addresses 0 operations issue)
' ====================================================================

Option Explicit

' Main function to update HeatMap status with detailed debugging
Sub UpdateHeatMapStatus()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim lastRowEval As Long
    Dim lastRowHeatMap As Long
    Dim i As Long
    Dim opCode As String
    Dim finalStatus As String
    Dim updatedCount As Long
    Dim startTime As Double
    Dim debugMsg As String
    Dim evalOpsFound As Long
    Dim heatMapOpsFound As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalOpsFound = 0
    heatMapOpsFound = 0
    debugMsg = ""
    
    ' Step 1: Verify Evaluation Results sheet exists
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = "✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    ' Step 2: Verify HeatMap Sheet exists
    Set wsHeatMap = Nothing
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found 'HeatMap Sheet'" & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing Evaluation Results..."
    
    ' Step 3: Find last rows
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = debugMsg & vbCrLf & "Sheet Dimensions:" & vbCrLf & _
               "- Evaluation Results: " & lastRowEval & " rows" & vbCrLf & _
               "- HeatMap Sheet: " & lastRowHeatMap & " rows" & vbCrLf
    
    ' Step 4: Find and process "Overall Status by Op Code" section
    Dim overallStartRow As Long
    overallStartRow = FindRowContaining(wsEval, "Overall Status by Op Code", 1, lastRowEval)
    
    If overallStartRow > 0 Then
        debugMsg = debugMsg & vbCrLf & "✓ Found 'Overall Status by Op Code' at row " & overallStartRow & vbCrLf
        
        ' Find header row (should be next row)
        Dim headerRow As Long
        headerRow = overallStartRow + 1
        
        ' Find columns in header
        Dim opCodeCol As Long, statusCol As Long
        opCodeCol = FindColumnInRow(wsEval, headerRow, "Op Code")
        statusCol = FindColumnInRow(wsEval, headerRow, "Final Status")
        
        If opCodeCol = 0 Then opCodeCol = 1 ' Default to column A
        If statusCol = 0 Then
            ' Try to find "Overall Status" instead
            statusCol = FindColumnInRow(wsEval, headerRow, "Overall Status")
        End If
        
        debugMsg = debugMsg & "  Op Code column: " & GetColumnLetter(opCodeCol) & " (" & opCodeCol & ")" & vbCrLf & _
                   "  Status column: " & GetColumnLetter(statusCol) & " (" & statusCol & ")" & vbCrLf
        
        ' Process operations in this section
        Application.StatusBar = "Processing Overall Status operations..."
        Dim dataStartRow As Long
        dataStartRow = headerRow + 1
        
        For i = dataStartRow To lastRowEval
            ' Check if we've hit the next section
            Dim cellValue As String
            cellValue = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            If InStr(1, cellValue, "Operation Mode Summary", vbTextCompare) > 0 Then
                debugMsg = debugMsg & "  Stopped at row " & i & " (found Operation Mode Summary)" & vbCrLf
                Exit For
            End If
            
            ' Check for empty rows (might indicate end of section)
            If cellValue = "" Then
                ' Check if next few rows are also empty
                If IsEndOfSection(wsEval, i, lastRowEval) Then
                    debugMsg = debugMsg & "  Stopped at row " & i & " (empty rows)" & vbCrLf
                    Exit For
                End If
                GoTo NextRow1
            End If
            
            ' Get operation code
            opCode = Trim(CStr(wsEval.Cells(i, opCodeCol).Value))
            
            ' Validate it's a numeric operation code
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 4 Then
                evalOpsFound = evalOpsFound + 1
                
                ' Get status
                If statusCol > 0 Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusCol).Value)))
                    
                    ' Update HeatMap if status is valid
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        If UpdateHeatMapOperation(wsHeatMap, opCode, finalStatus, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            End If
NextRow1:
        Next i
        
        debugMsg = debugMsg & "  Found " & evalOpsFound & " operations in Overall Status section" & vbCrLf
    Else
        debugMsg = debugMsg & vbCrLf & "⚠ WARNING: 'Overall Status by Op Code' section NOT found!" & vbCrLf
    End If
    
    ' Step 5: Find and process "Operation Mode Summary" section
    Dim summaryStartRow As Long
    summaryStartRow = FindRowContaining(wsEval, "Operation Mode Summary", 1, lastRowEval)
    
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & vbCrLf & "✓ Found 'Operation Mode Summary' at row " & summaryStartRow & vbCrLf
        
        ' Find header row
        Dim summaryHeaderRow As Long
        summaryHeaderRow = summaryStartRow + 1
        
        ' Find columns
        Dim summaryOpCodeCol As Long, summaryStatusCol As Long
        summaryOpCodeCol = FindColumnInRow(wsEval, summaryHeaderRow, "Op Code")
        summaryStatusCol = FindColumnInRow(wsEval, summaryHeaderRow, "Final Status")
        
        If summaryOpCodeCol = 0 Then summaryOpCodeCol = 1
        
        debugMsg = debugMsg & "  Op Code column: " & GetColumnLetter(summaryOpCodeCol) & " (" & summaryOpCodeCol & ")" & vbCrLf & _
                   "  Status column: " & GetColumnLetter(summaryStatusCol) & " (" & summaryStatusCol & ")" & vbCrLf
        
        ' Process operations
        Application.StatusBar = "Processing Operation Mode Summary..."
        Dim summaryDataRow As Long
        summaryDataRow = summaryHeaderRow + 1
        Dim summaryOpsFound As Long
        summaryOpsFound = 0
        
        For i = summaryDataRow To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, summaryOpCodeCol).Value))
            
            ' Stop at empty rows
            If opCode = "" Then
                If IsEndOfSection(wsEval, i, lastRowEval) Then
                    Exit For
                End If
                GoTo NextRow2
            End If
            
            ' Validate operation code
            If IsNumeric(opCode) And Len(opCode) >= 4 Then
                summaryOpsFound = summaryOpsFound + 1
                
                ' Get status
                If summaryStatusCol > 0 Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, summaryStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        If UpdateHeatMapOperation(wsHeatMap, opCode, finalStatus, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            End If
NextRow2:
        Next i
        
        debugMsg = debugMsg & "  Found " & summaryOpsFound & " operations in Summary section" & vbCrLf
        evalOpsFound = evalOpsFound + summaryOpsFound
    Else
        debugMsg = debugMsg & vbCrLf & "⚠ WARNING: 'Operation Mode Summary' section NOT found!" & vbCrLf
    End If
    
    ' Step 6: Count operations in HeatMap Sheet
    Application.StatusBar = "Analyzing HeatMap Sheet..."
    For i = 1 To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value)) ' Column A
        If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 4 Then
            heatMapOpsFound = heatMapOpsFound + 1
        End If
    Next i
    
    debugMsg = debugMsg & vbCrLf & "HeatMap Sheet: Found " & heatMapOpsFound & " operations in column A" & vbCrLf
    
    ' Final report
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    debugMsg = debugMsg & vbCrLf & "═══════════════════════════" & vbCrLf & _
               "RESULTS:" & vbCrLf & _
               "═══════════════════════════" & vbCrLf & _
               "Operations found in Evaluation: " & evalOpsFound & vbCrLf & _
               "Operations found in HeatMap: " & heatMapOpsFound & vbCrLf & _
               "Operations UPDATED: " & updatedCount & vbCrLf & _
               "Time elapsed: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf
    
    If updatedCount = 0 Then
        debugMsg = debugMsg & vbCrLf & "⚠ NO OPERATIONS WERE UPDATED!" & vbCrLf & vbCrLf & _
                   "Possible reasons:" & vbCrLf & _
                   "1. Operation codes don't match between sheets" & vbCrLf & _
                   "2. All statuses are N/A" & vbCrLf & _
                   "3. Status column not found correctly" & vbCrLf & _
                   "4. HeatMap Sheet structure different than expected"
    Else
        debugMsg = debugMsg & vbCrLf & "✓ Successfully updated " & updatedCount & " operations!"
    End If
    
    MsgBox debugMsg, IIf(updatedCount > 0, vbInformation, vbExclamation), "HeatMap Status Update - Debug Report"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "At line: " & Erl, _
           vbCritical, "Update Failed"
End Sub

' Helper: Update operation status in HeatMap Sheet
Private Function UpdateHeatMapOperation(ws As Worksheet, opCode As String, status As String, lastRow As Long) As Boolean
    Dim i As Long
    Dim heatMapOpCode As String
    Dim statusCol As Long
    
    UpdateHeatMapOperation = False
    
    ' Find status column (look for "Status" or "Current Status" in row 1 or nearby)
    statusCol = FindStatusColumn(ws, lastRow)
    If statusCol = 0 Then statusCol = 2 ' Default to column B if not found
    
    ' Search for operation code in column A
    For i = 1 To lastRow
        heatMapOpCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        ' Match operation code
        If heatMapOpCode = opCode Then
            ' Update status with colored dot
            ws.Cells(i, statusCol).Value = GetStatusDot(status)
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Size = 14
            ws.Cells(i, statusCol).Font.Color = GetStatusColor(status)
            
            UpdateHeatMapOperation = True
            Exit Function
        End If
    Next i
End Function

' Helper: Find status column in HeatMap Sheet
Private Function FindStatusColumn(ws As Worksheet, searchRows As Long) As Long
    Dim i As Long, j As Long
    Dim cellValue As String
    
    FindStatusColumn = 0
    
    ' Search first 10 rows for "Status" header
    For i = 1 To Application.Min(10, searchRows)
        For j = 1 To 20 ' Search first 20 columns
            cellValue = Trim(UCase(CStr(ws.Cells(i, j).Value)))
            If InStr(1, cellValue, "STATUS", vbTextCompare) > 0 Then
                FindStatusColumn = j
                Exit Function
            End If
        Next j
    Next i
End Function

' Helper: Find row containing specific text
Private Function FindRowContaining(ws As Worksheet, searchText As String, startRow As Long, endRow As Long) As Long
    Dim i As Long
    Dim cellValue As String
    
    FindRowContaining = 0
    
    For i = startRow To endRow
        cellValue = Trim(CStr(ws.Cells(i, 1).Value))
        If InStr(1, cellValue, searchText, vbTextCompare) > 0 Then
            FindRowContaining = i
            Exit Function
        End If
    Next i
End Function

' Helper: Find column by header text in a specific row
Private Function FindColumnInRow(ws As Worksheet, row As Long, headerText As String) As Long
    Dim j As Long
    Dim cellValue As String
    
    FindColumnInRow = 0
    
    For j = 1 To 30 ' Search first 30 columns
        cellValue = Trim(UCase(CStr(ws.Cells(row, j).Value)))
        If InStr(1, cellValue, UCase(headerText), vbTextCompare) > 0 Then
            FindColumnInRow = j
            Exit Function
        End If
    Next j
End Function

' Helper: Check if we've reached end of section (multiple empty rows)
Private Function IsEndOfSection(ws As Worksheet, startRow As Long, maxRow As Long) As Boolean
    Dim i As Long
    Dim emptyCount As Long
    
    IsEndOfSection = False
    emptyCount = 0
    
    ' Check next 3 rows
    For i = startRow To Application.Min(startRow + 2, maxRow)
        If Trim(CStr(ws.Cells(i, 1).Value)) = "" Then
            emptyCount = emptyCount + 1
        End If
    Next i
    
    ' If 2 or more empty rows, consider it end of section
    If emptyCount >= 2 Then
        IsEndOfSection = True
    End If
End Function

' Helper: Get status dot character
Private Function GetStatusDot(status As String) As String
    GetStatusDot = "●" ' Filled circle
End Function

' Helper: Get color for status
Private Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0) ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 192, 0) ' Yellow/Orange
        Case "GREEN"
            GetStatusColor = RGB(0, 176, 80) ' Green
        Case Else
            GetStatusColor = RGB(128, 128, 128) ' Gray for N/A
    End Select
End Function

' Helper: Get column letter from number
Private Function GetColumnLetter(colNum As Long) As String
    If colNum > 0 And colNum <= 16384 Then
        GetColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
    Else
        GetColumnLetter = "?"
    End If
End Function

' Helper: List all sheet names
Private Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Sheets
        sheetList = sheetList & ws.Name & ", "
    Next ws
    
    If Len(sheetList) > 2 Then
        sheetList = Left(sheetList, Len(sheetList) - 2)
    End If
    
    ListAllSheets = sheetList
End Function

' Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnExists As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Check if button already exists
    btnExists = False
    On Error Resume Next
    Set btn = ws.Buttons("UpdateHeatMapBtn")
    If Not btn Is Nothing Then btnExists = True
    On Error GoTo 0
    
    If btnExists Then
        MsgBox "Button already exists on HeatMap Sheet!", vbInformation
        Exit Sub
    End If
    
    ' Create button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = "UpdateHeatMapBtn"
    btn.Caption = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    btn.Font.Bold = True
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to transfer evaluation results.", vbInformation
End Sub
