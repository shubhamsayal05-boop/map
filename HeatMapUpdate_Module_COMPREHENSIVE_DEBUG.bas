Attribute VB_Name = "HeatMapUpdate_Debug"
' ====================================================================
' Module: HeatMapUpdate_Debug
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Version: Comprehensive Debug Edition
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
    Dim evalFound As Long, heatMapFound As Long
    Dim overallStartRow As Long, summaryStartRow As Long
    Dim statusColEval As Long, statusColSummary As Long
    Dim statusColHeatMap As Long
    Dim opCodeColSummary As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalFound = 0
    heatMapFound = 0
    debugMsg = ""
    
    ' ===== STEP 1: VERIFY SHEETS EXIST =====
    debugMsg = "=== DIAGNOSTIC REPORT ===" & vbCrLf & vbCrLf
    debugMsg = debugMsg & "STEP 1: Checking for required sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "  ✓ 'Evaluation Results' sheet found" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        ' Try alternative names
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
        If wsHeatMap Is Nothing Then
            Set wsHeatMap = ThisWorkbook.Sheets("Heat Map")
        End If
    End If
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: HeatMap sheet not found!" & vbCrLf & vbCrLf & _
               "Tried: 'HeatMap Sheet', 'HeatMap', 'Heat Map'" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "  ✓ '" & wsHeatMap.Name & "' sheet found" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing sheets..."
    
    ' ===== STEP 2: ANALYZE EVALUATION RESULTS STRUCTURE =====
    debugMsg = debugMsg & "STEP 2: Analyzing Evaluation Results structure..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "  Total rows: " & lastRowEval & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    overallStartRow = FindSectionRow(wsEval, "Overall Status by Op Code", lastRowEval)
    If overallStartRow > 0 Then
        debugMsg = debugMsg & "  ✓ 'Overall Status by Op Code' found at row " & overallStartRow & vbCrLf
        
        ' Find the header row (next row after section title)
        Dim headerRow As Long
        headerRow = overallStartRow + 1
        debugMsg = debugMsg & "    Header row: " & headerRow & vbCrLf
        debugMsg = debugMsg & "    Headers: "
        
        ' Show first 5 column headers
        For i = 1 To 5
            debugMsg = debugMsg & wsEval.Cells(headerRow, i).Value & " | "
        Next i
        debugMsg = debugMsg & "..." & vbCrLf
        
        ' Find "Final Status" or "Overall Status" column
        statusColEval = FindColumnByHeader(wsEval, headerRow, "Final Status")
        If statusColEval = 0 Then
            statusColEval = FindColumnByHeader(wsEval, headerRow, "Overall Status")
        End If
        
        If statusColEval > 0 Then
            debugMsg = debugMsg & "    ✓ Status column found at: " & ColumnLetter(statusColEval) & " (col " & statusColEval & ")" & vbCrLf
            
            ' Show sample data
            Dim sampleRow As Long
            sampleRow = headerRow + 1
            debugMsg = debugMsg & "    Sample: OpCode=" & wsEval.Cells(sampleRow, 1).Value & _
                       ", Status=" & wsEval.Cells(sampleRow, statusColEval).Value & vbCrLf
        Else
            debugMsg = debugMsg & "    ✗ Status column NOT found!" & vbCrLf
        End If
    Else
        debugMsg = debugMsg & "  ✗ 'Overall Status by Op Code' section NOT found!" & vbCrLf
    End If
    debugMsg = debugMsg & vbCrLf
    
    ' Find "Operation Mode Summary" section
    summaryStartRow = FindSectionRow(wsEval, "Operation Mode Summary", lastRowEval)
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "  ✓ 'Operation Mode Summary' found at row " & summaryStartRow & vbCrLf
        
        ' Find header row
        headerRow = summaryStartRow + 1
        debugMsg = debugMsg & "    Header row: " & headerRow & vbCrLf
        debugMsg = debugMsg & "    Headers: "
        
        For i = 1 To 5
            debugMsg = debugMsg & wsEval.Cells(headerRow, i).Value & " | "
        Next i
        debugMsg = debugMsg & "..." & vbCrLf
        
        ' Find Op Code and Final Status columns
        opCodeColSummary = FindColumnByHeader(wsEval, headerRow, "Op Code")
        statusColSummary = FindColumnByHeader(wsEval, headerRow, "Final Status")
        
        If opCodeColSummary > 0 Then
            debugMsg = debugMsg & "    ✓ Op Code column: " & ColumnLetter(opCodeColSummary) & vbCrLf
        End If
        If statusColSummary > 0 Then
            debugMsg = debugMsg & "    ✓ Status column: " & ColumnLetter(statusColSummary) & vbCrLf
        End If
    Else
        debugMsg = debugMsg & "  ✗ 'Operation Mode Summary' section NOT found!" & vbCrLf
    End If
    debugMsg = debugMsg & vbCrLf
    
    ' ===== STEP 3: ANALYZE HEATMAP STRUCTURE =====
    debugMsg = debugMsg & "STEP 3: Analyzing HeatMap structure..." & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "  Total rows: " & lastRowHeatMap & vbCrLf
    debugMsg = debugMsg & "  Column A header: " & wsHeatMap.Cells(1, 1).Value & vbCrLf
    
    ' Find "Status" column
    statusColHeatMap = FindStatusColumn(wsHeatMap)
    If statusColHeatMap > 0 Then
        debugMsg = debugMsg & "  ✓ Status column found at: " & ColumnLetter(statusColHeatMap) & " (col " & statusColHeatMap & ")" & vbCrLf
        debugMsg = debugMsg & "    Column header: " & wsHeatMap.Cells(1, statusColHeatMap).Value & vbCrLf
    Else
        debugMsg = debugMsg & "  ✗ Status column NOT found!" & vbCrLf
        debugMsg = debugMsg & "    Showing first 10 column headers:" & vbCrLf
        For i = 1 To 10
            debugMsg = debugMsg & "    Col " & i & " (" & ColumnLetter(i) & "): " & wsHeatMap.Cells(1, i).Value & vbCrLf
        Next i
    End If
    
    ' Show sample OpCodes from HeatMap
    debugMsg = debugMsg & "  Sample OpCodes (first 5):" & vbCrLf
    For i = 2 To Application.Min(6, lastRowHeatMap)
        debugMsg = debugMsg & "    Row " & i & ": " & wsHeatMap.Cells(i, 1).Value & vbCrLf
    Next i
    debugMsg = debugMsg & vbCrLf
    
    ' ===== STEP 4: ATTEMPT UPDATE =====
    debugMsg = debugMsg & "STEP 4: Attempting to update statuses..." & vbCrLf
    
    If overallStartRow > 0 And statusColEval > 0 And statusColHeatMap > 0 Then
        Application.StatusBar = "Processing Overall Status section..."
        
        For i = overallStartRow + 2 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Stop if we hit next section
            If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                evalFound = evalFound + 1
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColEval).Value)))
                
                If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                    If UpdateOperationStatusDirect(wsHeatMap, opCode, finalStatus, statusColHeatMap, lastRowHeatMap) Then
                        updatedCount = updatedCount + 1
                    End If
                End If
            End If
        Next i
        
        debugMsg = debugMsg & "  Sub-operations: Found " & evalFound & ", Updated " & updatedCount & vbCrLf
    Else
        debugMsg = debugMsg & "  ✗ Cannot update sub-operations (missing data)" & vbCrLf
    End If
    
    ' Process Operation Mode Summary
    If summaryStartRow > 0 And opCodeColSummary > 0 And statusColSummary > 0 And statusColHeatMap > 0 Then
        Application.StatusBar = "Processing Operation Mode Summary..."
        
        Dim summaryFound As Long, summaryUpdated As Long
        summaryFound = 0
        summaryUpdated = 0
        
        For i = summaryStartRow + 2 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, opCodeColSummary).Value))
            
            If opCode = "" Then Exit For ' End of section
            
            If IsNumeric(opCode) And Len(opCode) = 8 Then
                summaryFound = summaryFound + 1
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColSummary).Value)))
                
                If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                    If UpdateOperationStatusDirect(wsHeatMap, opCode, finalStatus, statusColHeatMap, lastRowHeatMap) Then
                        summaryUpdated = summaryUpdated + 1
                        updatedCount = updatedCount + 1
                    End If
                End If
            End If
        Next i
        
        debugMsg = debugMsg & "  Parent operations: Found " & summaryFound & ", Updated " & summaryUpdated & vbCrLf
    Else
        debugMsg = debugMsg & "  ✗ Cannot update parent operations (missing data)" & vbCrLf
    End If
    
    debugMsg = debugMsg & vbCrLf
    debugMsg = debugMsg & "=== SUMMARY ===" & vbCrLf
    debugMsg = debugMsg & "Total operations updated: " & updatedCount & vbCrLf
    debugMsg = debugMsg & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show results
    If updatedCount > 0 Then
        MsgBox debugMsg & vbCrLf & vbCrLf & _
               "✓ Successfully updated " & updatedCount & " operation(s)!", _
               vbInformation, "Update Complete"
    Else
        MsgBox debugMsg & vbCrLf & vbCrLf & _
               "⚠ No operations were updated." & vbCrLf & _
               "Please review the diagnostic information above.", _
               vbExclamation, "No Updates Made"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error occurred: " & Err.Description & vbCrLf & vbCrLf & _
           "Diagnostic info:" & vbCrLf & debugMsg, vbCritical, "Error"
End Sub

' Helper function to find a section by name
Private Function FindSectionRow(ws As Worksheet, sectionName As String, lastRow As Long) As Long
    Dim i As Long
    FindSectionRow = 0
    
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, 1).Value), sectionName, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find column by header name
Private Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim i As Long
    Dim cellValue As String
    FindColumnByHeader = 0
    
    For i = 1 To 20 ' Check first 20 columns
        cellValue = Trim(CStr(ws.Cells(headerRow, i).Value))
        If InStr(1, cellValue, headerName, vbTextCompare) > 0 Then
            FindColumnByHeader = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find Status column in HeatMap
Private Function FindStatusColumn(ws As Worksheet) As Long
    Dim i As Long
    Dim cellValue As String
    FindStatusColumn = 0
    
    ' Check first row for "Status" header
    For i = 1 To 30
        cellValue = Trim(CStr(ws.Cells(1, i).Value))
        If InStr(1, cellValue, "Status", vbTextCompare) > 0 Then
            ' Make sure it's not "Current Status" or other variants
            ' Prefer exact match or "Status P1"
            If cellValue = "Status" Or InStr(1, cellValue, "Status P1", vbTextCompare) > 0 Or _
               InStr(1, cellValue, "Current Status", vbTextCompare) > 0 Then
                FindStatusColumn = i
                Exit Function
            End If
        End If
    Next i
End Function

' Update operation status directly with known column
Private Function UpdateOperationStatusDirect(ws As Worksheet, opCode As String, status As String, _
                                            statusCol As Long, lastRow As Long) As Boolean
    Dim i As Long
    Dim wsOpCode As String
    
    UpdateOperationStatusDirect = False
    
    For i = 2 To lastRow
        wsOpCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If wsOpCode = opCode Then
            ' Found match - update status with colored dot
            ws.Cells(i, statusCol).Value = GetStatusSymbol(status)
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Size = 12
            ws.Cells(i, statusCol).Font.Color = GetStatusColor(status)
            ws.Cells(i, statusCol).HorizontalAlignment = xlCenter
            
            UpdateOperationStatusDirect = True
            Exit Function
        End If
    Next i
End Function

' Get status symbol (filled circle)
Private Function GetStatusSymbol(status As String) As String
    GetStatusSymbol = "l" ' Wingdings filled circle
End Function

' Get color for status
Private Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0)     ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 192, 0)   ' Yellow/Orange
        Case "GREEN"
            GetStatusColor = RGB(0, 176, 80)    ' Green
        Case Else
            GetStatusColor = RGB(128, 128, 128) ' Gray for N/A
    End Select
End Function

' Get column letter from number
Private Function ColumnLetter(colNum As Long) As String
    Dim n As Long
    n = colNum
    ColumnLetter = ""
    
    Do While n > 0
        ColumnLetter = Chr((n - 1) Mod 26 + 65) & ColumnLetter
        n = (n - 1) \ 26
    Loop
End Function

' List all sheets in workbook
Private Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim result As String
    result = ""
    
    For Each ws In ThisWorkbook.Worksheets
        result = result & "  - " & ws.Name & vbCrLf
    Next ws
    
    ListAllSheets = result
End Function

' Create button on HeatMap sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("HeatMap")
    End If
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("Heat Map")
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if present
    On Error Resume Next
    ws.Buttons("UpdateHeatMapBtn").Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = "UpdateHeatMapBtn"
    btn.Caption = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on '" & ws.Name & "' sheet!" & vbCrLf & vbCrLf & _
           "Click the button after running evaluation to transfer statuses.", vbInformation
End Sub
