Attribute VB_Name = "HeatMapUpdateDiagnosticV2"
' ====================================================================
' Module: HeatMapUpdateDiagnosticV2
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive diagnostics
' Author: GitHub Copilot
' Version: 2.0 - Enhanced diagnostic version
' ====================================================================

Option Explicit

' Main function to update HeatMap status with full diagnostic output
Sub UpdateHeatMapStatusWithDiagnostics()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim lastRowEval As Long
    Dim lastRowHeatMap As Long
    Dim i As Long, j As Long
    Dim opCode As String
    Dim finalStatus As String
    Dim updatedCount As Long
    Dim startTime As Double
    Dim diagnosticReport As String
    Dim evalOpCodes As Long, heatMapOpCodes As Long
    Dim statusCol As Long
    Dim overallStartRow As Long, summaryStartRow As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalOpCodes = 0
    heatMapOpCodes = 0
    diagnosticReport = "==== HEATMAP UPDATE DIAGNOSTIC REPORT ====" & vbCrLf & vbCrLf
    
    ' ========== STEP 1: Verify Sheets Exist ==========
    diagnosticReport = diagnosticReport & "STEP 1: Checking for required sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        diagnosticReport = diagnosticReport & "  ❌ ERROR: 'Evaluation Results' sheet NOT FOUND!" & vbCrLf
        diagnosticReport = diagnosticReport & "  Available sheets: " & ListAllSheets() & vbCrLf
        MsgBox diagnosticReport, vbCritical, "Diagnostic Report"
        Exit Sub
    End If
    diagnosticReport = diagnosticReport & "  ✓ 'Evaluation Results' sheet found" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        ' Try alternate names
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
        If wsHeatMap Is Nothing Then
            Set wsHeatMap = ThisWorkbook.Sheets("Heatmap Sheet")
        End If
    End If
    
    If wsHeatMap Is Nothing Then
        diagnosticReport = diagnosticReport & "  ❌ ERROR: 'HeatMap Sheet' NOT FOUND!" & vbCrLf
        diagnosticReport = diagnosticReport & "  Available sheets: " & ListAllSheets() & vbCrLf
        MsgBox diagnosticReport, vbCritical, "Diagnostic Report"
        Exit Sub
    End If
    diagnosticReport = diagnosticReport & "  ✓ '" & wsHeatMap.Name & "' sheet found" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' ========== STEP 2: Analyze Evaluation Results Sheet Structure ==========
    diagnosticReport = diagnosticReport & "STEP 2: Analyzing 'Evaluation Results' sheet..." & vbCrLf
    
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    diagnosticReport = diagnosticReport & "  Last row with data: " & lastRowEval & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    overallStartRow = FindSectionRow(wsEval, "Overall Status by Op Code", lastRowEval)
    If overallStartRow > 0 Then
        diagnosticReport = diagnosticReport & "  ✓ 'Overall Status by Op Code' section found at row " & overallStartRow & vbCrLf
        
        ' Scan first 20 rows of this section to show structure
        diagnosticReport = diagnosticReport & "  First few rows of data:" & vbCrLf
        For i = overallStartRow To Application.Min(overallStartRow + 5, lastRowEval)
            diagnosticReport = diagnosticReport & "    Row " & i & ": [A]=" & Left(wsEval.Cells(i, 1).Value, 20) & _
                              " [B]=" & Left(wsEval.Cells(i, 2).Value, 20) & _
                              " [C]=" & Left(wsEval.Cells(i, 3).Value, 15) & vbCrLf
        Next i
        
        ' Find Final Status column
        statusCol = FindColumnByHeader(wsEval, overallStartRow + 1, "Final Status")
        If statusCol > 0 Then
            diagnosticReport = diagnosticReport & "  ✓ 'Final Status' column found at column " & statusCol & " (" & ColumnLetter(statusCol) & ")" & vbCrLf
        Else
            diagnosticReport = diagnosticReport & "  ❌ 'Final Status' column NOT FOUND in header row" & vbCrLf
            diagnosticReport = diagnosticReport & "  Header row " & (overallStartRow + 1) & " contents:" & vbCrLf
            For j = 1 To 15
                If wsEval.Cells(overallStartRow + 1, j).Value <> "" Then
                    diagnosticReport = diagnosticReport & "    Col " & ColumnLetter(j) & ": " & wsEval.Cells(overallStartRow + 1, j).Value & vbCrLf
                End If
            Next j
        End If
        
        ' Count operation codes in this section
        For i = overallStartRow + 2 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                evalOpCodes = evalOpCodes + 1
            End If
        Next i
        diagnosticReport = diagnosticReport & "  Operation codes found in section: " & evalOpCodes & vbCrLf
    Else
        diagnosticReport = diagnosticReport & "  ❌ 'Overall Status by Op Code' section NOT FOUND!" & vbCrLf
    End If
    
    ' Find "Operation Mode Summary" section
    summaryStartRow = FindSectionRow(wsEval, "Operation Mode Summary", lastRowEval)
    If summaryStartRow > 0 Then
        diagnosticReport = diagnosticReport & "  ✓ 'Operation Mode Summary' section found at row " & summaryStartRow & vbCrLf
    Else
        diagnosticReport = diagnosticReport & "  ⚠ 'Operation Mode Summary' section NOT FOUND" & vbCrLf
    End If
    
    diagnosticReport = diagnosticReport & vbCrLf
    
    ' ========== STEP 3: Analyze HeatMap Sheet Structure ==========
    diagnosticReport = diagnosticReport & "STEP 3: Analyzing '" & wsHeatMap.Name & "' sheet..." & vbCrLf
    
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    diagnosticReport = diagnosticReport & "  Last row with data: " & lastRowHeatMap & vbCrLf
    
    ' Show first few rows
    diagnosticReport = diagnosticReport & "  First few rows:" & vbCrLf
    For i = 1 To Application.Min(5, lastRowHeatMap)
        diagnosticReport = diagnosticReport & "    Row " & i & ": [A]=" & Left(wsHeatMap.Cells(i, 1).Value, 20) & _
                          " [B]=" & Left(wsHeatMap.Cells(i, 2).Value, 20) & vbCrLf
    Next i
    
    ' Find Status column
    Dim heatMapStatusCol As Long
    heatMapStatusCol = FindColumnByHeader(wsHeatMap, 1, "Status")
    If heatMapStatusCol = 0 Then
        ' Try alternate names
        heatMapStatusCol = FindColumnByHeader(wsHeatMap, 1, "Current Status")
        If heatMapStatusCol = 0 Then
            heatMapStatusCol = FindColumnByHeader(wsHeatMap, 1, "Current Status P1")
        End If
    End If
    
    If heatMapStatusCol > 0 Then
        diagnosticReport = diagnosticReport & "  ✓ Status column found at column " & heatMapStatusCol & " (" & ColumnLetter(heatMapStatusCol) & ")" & vbCrLf
    Else
        diagnosticReport = diagnosticReport & "  ❌ Status column NOT FOUND" & vbCrLf
        diagnosticReport = diagnosticReport & "  Header row contents:" & vbCrLf
        For j = 1 To 10
            If wsHeatMap.Cells(1, j).Value <> "" Then
                diagnosticReport = diagnosticReport & "    Col " & ColumnLetter(j) & ": " & wsHeatMap.Cells(1, j).Value & vbCrLf
            End If
        Next j
    End If
    
    ' Count operation codes in HeatMap
    For i = 2 To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
            heatMapOpCodes = heatMapOpCodes + 1
        End If
    Next i
    diagnosticReport = diagnosticReport & "  Operation codes found: " & heatMapOpCodes & vbCrLf
    
    diagnosticReport = diagnosticReport & vbCrLf
    
    ' ========== STEP 4: Attempt Update (if possible) ==========
    diagnosticReport = diagnosticReport & "STEP 4: Attempting status update..." & vbCrLf
    
    If overallStartRow > 0 And statusCol > 0 And heatMapStatusCol > 0 Then
        diagnosticReport = diagnosticReport & "  ✓ All required elements found. Proceeding with update..." & vbCrLf
        
        ' Update statuses
        For i = overallStartRow + 2 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Stop if we hit next section
            If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusCol).Value)))
                
                If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                    ' Find in HeatMap and update
                    For j = 2 To lastRowHeatMap
                        If Trim(CStr(wsHeatMap.Cells(j, 1).Value)) = opCode Then
                            ' Update status with colored dot
                            wsHeatMap.Cells(j, heatMapStatusCol).Value = "●"
                            wsHeatMap.Cells(j, heatMapStatusCol).Font.Name = "Wingdings"
                            
                            Select Case finalStatus
                                Case "RED"
                                    wsHeatMap.Cells(j, heatMapStatusCol).Font.Color = RGB(255, 0, 0)
                                Case "YELLOW"
                                    wsHeatMap.Cells(j, heatMapStatusCol).Font.Color = RGB(255, 192, 0)
                                Case "GREEN"
                                    wsHeatMap.Cells(j, heatMapStatusCol).Font.Color = RGB(0, 176, 80)
                                Case Else
                                    wsHeatMap.Cells(j, heatMapStatusCol).Font.Color = RGB(128, 128, 128)
                            End Select
                            
                            updatedCount = updatedCount + 1
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i
        
        diagnosticReport = diagnosticReport & "  ✓ Update completed: " & updatedCount & " operations updated" & vbCrLf
    Else
        diagnosticReport = diagnosticReport & "  ❌ Cannot proceed with update - missing required elements:" & vbCrLf
        If overallStartRow = 0 Then diagnosticReport = diagnosticReport & "    - Overall Status section not found" & vbCrLf
        If statusCol = 0 Then diagnosticReport = diagnosticReport & "    - Final Status column not found" & vbCrLf
        If heatMapStatusCol = 0 Then diagnosticReport = diagnosticReport & "    - HeatMap Status column not found" & vbCrLf
    End If
    
    diagnosticReport = diagnosticReport & vbCrLf
    
    ' ========== STEP 5: Summary ==========
    diagnosticReport = diagnosticReport & "===== SUMMARY =====" & vbCrLf
    diagnosticReport = diagnosticReport & "Operation codes in Evaluation Results: " & evalOpCodes & vbCrLf
    diagnosticReport = diagnosticReport & "Operation codes in HeatMap: " & heatMapOpCodes & vbCrLf
    diagnosticReport = diagnosticReport & "Statuses updated: " & updatedCount & vbCrLf
    diagnosticReport = diagnosticReport & "Time elapsed: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show full diagnostic report
    MsgBox diagnosticReport, vbInformation, "HeatMap Update Diagnostic Report"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    diagnosticReport = diagnosticReport & vbCrLf & "ERROR: " & Err.Description & " (Line: " & Erl & ")"
    MsgBox diagnosticReport, vbCritical, "Error in Update"
End Sub

' Helper function to find a section by title
Private Function FindSectionRow(ws As Worksheet, sectionTitle As String, lastRow As Long) As Long
    Dim i As Long
    FindSectionRow = 0
    
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, 1).Value), sectionTitle, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find column by header name
Private Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim j As Long
    FindColumnByHeader = 0
    
    For j = 1 To 50 ' Search first 50 columns
        If InStr(1, CStr(ws.Cells(headerRow, j).Value), headerName, vbTextCompare) > 0 Then
            FindColumnByHeader = j
            Exit Function
        End If
    Next j
End Function

' Helper function to convert column number to letter
Private Function ColumnLetter(colNum As Long) As String
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

' Helper function to list all sheet names
Private Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        sheetList = sheetList & ws.Name & ", "
    Next ws
    
    If Len(sheetList) > 2 Then
        sheetList = Left(sheetList, Len(sheetList) - 2)
    End If
    
    ListAllSheets = sheetList
End Function

' Create button to run the diagnostic update
Sub CreateDiagnosticUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("HeatMap")
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets("Heatmap Sheet")
        End If
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Cannot find HeatMap Sheet to add button!", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing button if present
    On Error Resume Next
    ws.Buttons("UpdateHeatMapDiagnostic").Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    btn.Name = "UpdateHeatMapDiagnostic"
    btn.Text = "Update HeatMap (Diagnostic)"
    btn.OnAction = "UpdateHeatMapStatusWithDiagnostics"
    
    MsgBox "Diagnostic button created on '" & ws.Name & "' sheet!" & vbCrLf & vbCrLf & _
           "Click the button to run the diagnostic update.", vbInformation
End Sub
