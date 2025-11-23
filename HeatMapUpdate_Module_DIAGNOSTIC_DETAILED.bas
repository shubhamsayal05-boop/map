Attribute VB_Name = "HeatMapUpdate_Diagnostic"
' ====================================================================
' Module: HeatMapUpdate_Diagnostic_Detailed
' Purpose: Transfer evaluation results to HeatMap Sheet with detailed diagnostics
' ====================================================================

Option Explicit

' Main function to update HeatMap status with detailed diagnostic output
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
    Dim diagnosticMsg As String
    Dim evalOpCodes As String, heatMapOpCodes As String
    Dim statusColumn As Long
    Dim overallStartRow As Long, summaryStartRow As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    diagnosticMsg = "=== HEATMAP UPDATE DIAGNOSTIC REPORT ===" & vbCrLf & vbCrLf
    
    ' Step 1: Check if sheets exist
    diagnosticMsg = diagnosticMsg & "STEP 1: Checking Sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets in workbook:" & vbCrLf & GetAllSheetNames(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    diagnosticMsg = diagnosticMsg & "  ✓ 'Evaluation Results' sheet found" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        ' Try alternate names
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
        If wsHeatMap Is Nothing Then
            Set wsHeatMap = ThisWorkbook.Sheets("Heatmap Sheet")
        End If
    End If
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: HeatMap sheet not found!" & vbCrLf & vbCrLf & _
               "Looking for: 'HeatMap Sheet', 'HeatMap', or 'Heatmap Sheet'" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetAllSheetNames(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    diagnosticMsg = diagnosticMsg & "  ✓ HeatMap sheet found: '" & wsHeatMap.Name & "'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    ' Step 2: Analyze Evaluation Results structure
    diagnosticMsg = diagnosticMsg & "STEP 2: Analyzing Evaluation Results Sheet..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    diagnosticMsg = diagnosticMsg & "  Total rows: " & lastRowEval & vbCrLf
    
    ' Find sections
    overallStartRow = FindSectionRow(wsEval, "Overall Status by Op Code", lastRowEval)
    summaryStartRow = FindSectionRow(wsEval, "Operation Mode Summary", lastRowEval)
    
    If overallStartRow > 0 Then
        diagnosticMsg = diagnosticMsg & "  ✓ 'Overall Status by Op Code' section found at row " & overallStartRow & vbCrLf
        
        ' Show first few op codes from this section
        diagnosticMsg = diagnosticMsg & "    Sample Op Codes from section:" & vbCrLf
        Dim sampleCount As Long
        sampleCount = 0
        For i = overallStartRow + 2 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            If opCode <> "" And IsNumeric(opCode) And sampleCount < 5 Then
                diagnosticMsg = diagnosticMsg & "      Row " & i & ": " & opCode & vbCrLf
                sampleCount = sampleCount + 1
            End If
            If sampleCount >= 5 Then Exit For
        Next i
    Else
        diagnosticMsg = diagnosticMsg & "  ✗ 'Overall Status by Op Code' section NOT found" & vbCrLf
    End If
    
    If summaryStartRow > 0 Then
        diagnosticMsg = diagnosticMsg & "  ✓ 'Operation Mode Summary' section found at row " & summaryStartRow & vbCrLf
    Else
        diagnosticMsg = diagnosticMsg & "  ✗ 'Operation Mode Summary' section NOT found" & vbCrLf
    End If
    diagnosticMsg = diagnosticMsg & vbCrLf
    
    ' Step 3: Analyze HeatMap Sheet structure
    diagnosticMsg = diagnosticMsg & "STEP 3: Analyzing HeatMap Sheet..." & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    diagnosticMsg = diagnosticMsg & "  Total rows: " & lastRowHeatMap & vbCrLf
    
    ' Find Status column
    statusColumn = FindStatusColumn(wsHeatMap)
    If statusColumn > 0 Then
        diagnosticMsg = diagnosticMsg & "  ✓ Status column found at column " & statusColumn & " (" & ColumnLetter(statusColumn) & ")" & vbCrLf
    Else
        diagnosticMsg = diagnosticMsg & "  ✗ Status column NOT found (looking for 'Status' or 'Current Status')" & vbCrLf
    End If
    
    ' Show first few op codes from HeatMap
    diagnosticMsg = diagnosticMsg & "  Sample Op Codes from HeatMap:" & vbCrLf
    sampleCount = 0
    For i = 2 To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" And IsNumeric(opCode) And sampleCount < 5 Then
            diagnosticMsg = diagnosticMsg & "    Row " & i & ": " & opCode & vbCrLf
            sampleCount = sampleCount + 1
        End If
        If sampleCount >= 5 Then Exit For
    Next i
    diagnosticMsg = diagnosticMsg & vbCrLf
    
    ' Step 4: Check for column header details
    diagnosticMsg = diagnosticMsg & "STEP 4: Column Headers..." & vbCrLf
    diagnosticMsg = diagnosticMsg & "  HeatMap Sheet Row 1:" & vbCrLf
    For i = 1 To 10 ' First 10 columns
        Dim headerVal As String
        headerVal = Trim(CStr(wsHeatMap.Cells(1, i).Value))
        If headerVal <> "" Then
            diagnosticMsg = diagnosticMsg & "    Column " & ColumnLetter(i) & ": '" & headerVal & "'" & vbCrLf
        End If
    Next i
    
    If overallStartRow > 0 Then
        diagnosticMsg = diagnosticMsg & "  Evaluation Results Row " & (overallStartRow + 1) & ":" & vbCrLf
        For i = 1 To 15 ' First 15 columns
            headerVal = Trim(CStr(wsEval.Cells(overallStartRow + 1, i).Value))
            If headerVal <> "" Then
                diagnosticMsg = diagnosticMsg & "    Column " & ColumnLetter(i) & ": '" & headerVal & "'" & vbCrLf
            End If
        Next i
    End If
    diagnosticMsg = diagnosticMsg & vbCrLf
    
    ' Step 5: Attempt to match and update
    diagnosticMsg = diagnosticMsg & "STEP 5: Attempting to Match and Update..." & vbCrLf
    
    If overallStartRow = 0 Then
        diagnosticMsg = diagnosticMsg & "  ✗ Cannot proceed: 'Overall Status by Op Code' section not found" & vbCrLf
    ElseIf statusColumn = 0 Then
        diagnosticMsg = diagnosticMsg & "  ✗ Cannot proceed: Status column not found in HeatMap" & vbCrLf
    Else
        ' Find Final Status column
        Dim finalStatusCol As Long
        finalStatusCol = FindColumnByHeader(wsEval, overallStartRow + 1, "Final Status")
        
        If finalStatusCol = 0 Then
            diagnosticMsg = diagnosticMsg & "  ✗ 'Final Status' column not found in Evaluation Results" & vbCrLf
        Else
            diagnosticMsg = diagnosticMsg & "  ✓ 'Final Status' column found at column " & finalStatusCol & vbCrLf
            diagnosticMsg = diagnosticMsg & "  Processing updates..." & vbCrLf & vbCrLf
            
            ' Now perform the actual update
            Dim matchCount As Long
            matchCount = 0
            
            For i = overallStartRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                ' Stop at next section
                If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
                    Exit For
                End If
                
                If opCode <> "" And IsNumeric(opCode) Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, finalStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "FINAL STATUS" Then
                        ' Find in HeatMap
                        For j = 2 To lastRowHeatMap
                            If Trim(CStr(wsHeatMap.Cells(j, 1).Value)) = opCode Then
                                matchCount = matchCount + 1
                                ' Update status
                                wsHeatMap.Cells(j, statusColumn).Value = finalStatus
                                
                                ' Apply color
                                Select Case finalStatus
                                    Case "RED"
                                        wsHeatMap.Cells(j, statusColumn).Interior.Color = RGB(255, 0, 0)
                                        wsHeatMap.Cells(j, statusColumn).Font.Color = RGB(255, 255, 255)
                                    Case "YELLOW"
                                        wsHeatMap.Cells(j, statusColumn).Interior.Color = RGB(255, 255, 0)
                                        wsHeatMap.Cells(j, statusColumn).Font.Color = RGB(0, 0, 0)
                                    Case "GREEN"
                                        wsHeatMap.Cells(j, statusColumn).Interior.Color = RGB(0, 255, 0)
                                        wsHeatMap.Cells(j, statusColumn).Font.Color = RGB(0, 0, 0)
                                    Case Else
                                        wsHeatMap.Cells(j, statusColumn).Interior.Color = RGB(192, 192, 192)
                                        wsHeatMap.Cells(j, statusColumn).Font.Color = RGB(0, 0, 0)
                                End Select
                                
                                updatedCount = updatedCount + 1
                                
                                ' Log first 5 updates
                                If updatedCount <= 5 Then
                                    diagnosticMsg = diagnosticMsg & "    Update #" & updatedCount & ": Op Code " & opCode & " → " & finalStatus & vbCrLf
                                End If
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Next i
            
            If updatedCount > 5 Then
                diagnosticMsg = diagnosticMsg & "    ... and " & (updatedCount - 5) & " more updates" & vbCrLf
            End If
        End If
    End If
    
    diagnosticMsg = diagnosticMsg & vbCrLf & "=== SUMMARY ===" & vbCrLf
    diagnosticMsg = diagnosticMsg & "Operations matched and updated: " & updatedCount & vbCrLf
    diagnosticMsg = diagnosticMsg & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show diagnostic report
    MsgBox diagnosticMsg, vbInformation, "HeatMap Update Diagnostic Report"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error occurred: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl, vbCritical, "Error"
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
    Dim i As Long
    FindColumnByHeader = 0
    
    For i = 1 To 50 ' Check first 50 columns
        If InStr(1, CStr(ws.Cells(headerRow, i).Value), headerName, vbTextCompare) > 0 Then
            FindColumnByHeader = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find Status column in HeatMap
Private Function FindStatusColumn(ws As Worksheet) As Long
    Dim i As Long
    Dim headerVal As String
    FindStatusColumn = 0
    
    For i = 1 To 20 ' Check first 20 columns
        headerVal = Trim(UCase(CStr(ws.Cells(1, i).Value)))
        If headerVal = "STATUS" Or headerVal = "CURRENT STATUS" Or _
           InStr(1, headerVal, "STATUS", vbTextCompare) > 0 Then
            FindStatusColumn = i
            Exit Function
        End If
    Next i
End Function

' Helper function to get all sheet names
Private Function GetAllSheetNames() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        sheetList = sheetList & "  - " & ws.Name & vbCrLf
    Next ws
    
    GetAllSheetNames = sheetList
End Function

' Helper function to convert column number to letter
Private Function ColumnLetter(colNum As Long) As String
    Dim result As String
    Dim num As Long
    
    num = colNum
    Do While num > 0
        result = Chr(65 + ((num - 1) Mod 26)) & result
        num = (num - 1) \ 26
    Loop
    
    ColumnLetter = result
End Function

' Button creation function
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
        Set ws = ThisWorkbook.Sheets("Heatmap Sheet")
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap sheet not found! Cannot create button.", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if present
    On Error Resume Next
    ws.Buttons("Update HeatMap Status").Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    btn.Name = "Update HeatMap Status"
    btn.Text = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on '" & ws.Name & "' sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update status with detailed diagnostics.", vbInformation
End Sub
