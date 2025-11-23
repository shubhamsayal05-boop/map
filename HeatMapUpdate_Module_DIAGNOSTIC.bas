Attribute VB_Name = "HeatMapUpdate_Diagnostic"
' ====================================================================
' Module: HeatMapUpdate_Diagnostic
' Purpose: Transfer evaluation results to HeatMap Sheet with detailed diagnostics
' ====================================================================

Option Explicit

' Main function to update HeatMap status with comprehensive debugging
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
    Dim debugReport As String
    Dim evalOpCodes As Long, heatMapOpCodes As Long
    Dim matchedCodes As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalOpCodes = 0
    heatMapOpCodes = 0
    matchedCodes = 0
    debugReport = "=== HEATMAP UPDATE DIAGNOSTIC REPORT ===" & vbCrLf & vbCrLf
    
    ' Step 1: Verify sheets exist
    debugReport = debugReport & "STEP 1: Checking for required sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: Cannot find 'Evaluation Results' sheet!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugReport = debugReport & "✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        ' Try alternative names
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
        If wsHeatMap Is Nothing Then
            Set wsHeatMap = ThisWorkbook.Sheets("Heat Map")
        End If
    End If
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: Cannot find 'HeatMap Sheet' (or 'HeatMap' or 'Heat Map')!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugReport = debugReport & "✓ Found HeatMap sheet: '" & wsHeatMap.Name & "'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' Step 2: Analyze Evaluation Results sheet structure
    debugReport = debugReport & "STEP 2: Analyzing Evaluation Results sheet..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugReport = debugReport & "Last row with data: " & lastRowEval & vbCrLf
    
    ' Look for key sections
    Dim overallRow As Long, summaryRow As Long
    overallRow = FindRowContaining(wsEval, "Overall Status by Op Code", lastRowEval)
    summaryRow = FindRowContaining(wsEval, "Operation Mode Summary", lastRowEval)
    
    debugReport = debugReport & "Found sections:" & vbCrLf
    If overallRow > 0 Then
        debugReport = debugReport & "  - 'Overall Status by Op Code' at row " & overallRow & vbCrLf
    Else
        debugReport = debugReport & "  - 'Overall Status by Op Code' NOT FOUND" & vbCrLf
    End If
    
    If summaryRow > 0 Then
        debugReport = debugReport & "  - 'Operation Mode Summary' at row " & summaryRow & vbCrLf
    Else
        debugReport = debugReport & "  - 'Operation Mode Summary' NOT FOUND" & vbCrLf
    End If
    debugReport = debugReport & vbCrLf
    
    ' Step 3: Find column headers and structure
    debugReport = debugReport & "STEP 3: Finding column structure..." & vbCrLf
    
    Dim statusCol As Long, opCodeCol As Long
    Dim headerRow As Long
    
    ' If we found "Overall Status by Op Code", the headers are likely in the next row
    If overallRow > 0 Then
        headerRow = overallRow + 1
        debugReport = debugReport & "Looking for headers in row " & headerRow & ":" & vbCrLf
        debugReport = debugReport & "  Column A: '" & wsEval.Cells(headerRow, 1).Value & "'" & vbCrLf
        debugReport = debugReport & "  Column B: '" & wsEval.Cells(headerRow, 2).Value & "'" & vbCrLf
        debugReport = debugReport & "  Column C: '" & wsEval.Cells(headerRow, 3).Value & "'" & vbCrLf
        debugReport = debugReport & "  Column D: '" & wsEval.Cells(headerRow, 4).Value & "'" & vbCrLf
        debugReport = debugReport & "  Column E: '" & wsEval.Cells(headerRow, 5).Value & "'" & vbCrLf
        
        ' Find "Final Status" or "Overall Status" column
        statusCol = FindColumnInRow(wsEval, headerRow, Array("Final Status", "Overall Status", "Status"))
        If statusCol > 0 Then
            debugReport = debugReport & "✓ Found status column at position " & statusCol & vbCrLf
        Else
            debugReport = debugReport & "✗ Could not find status column!" & vbCrLf
        End If
        debugReport = debugReport & vbCrLf
    End If
    
    ' Step 4: Count operation codes in Evaluation Results
    debugReport = debugReport & "STEP 4: Scanning for operation codes in Evaluation Results..." & vbCrLf
    
    If overallRow > 0 And statusCol > 0 Then
        Dim dataStartRow As Long
        dataStartRow = headerRow + 1
        
        For i = dataStartRow To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Stop at next section
            If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 8 Then
                evalOpCodes = evalOpCodes + 1
                If evalOpCodes <= 5 Then
                    debugReport = debugReport & "  Sample " & evalOpCodes & ": Op Code=" & opCode & _
                                  ", Status=" & wsEval.Cells(i, statusCol).Value & vbCrLf
                End If
            End If
        Next i
        debugReport = debugReport & "Total operation codes found: " & evalOpCodes & vbCrLf & vbCrLf
    End If
    
    ' Step 5: Analyze HeatMap Sheet
    debugReport = debugReport & "STEP 5: Analyzing HeatMap Sheet..." & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugReport = debugReport & "Last row with data: " & lastRowHeatMap & vbCrLf
    
    ' Check first few rows for structure
    debugReport = debugReport & "First row content:" & vbCrLf
    debugReport = debugReport & "  Column A (row 1): '" & wsHeatMap.Cells(1, 1).Value & "'" & vbCrLf
    debugReport = debugReport & "  Column B (row 1): '" & wsHeatMap.Cells(1, 2).Value & "'" & vbCrLf
    debugReport = debugReport & "  Column C (row 1): '" & wsHeatMap.Cells(1, 3).Value & "'" & vbCrLf
    
    ' Find "Status" column in HeatMap
    Dim heatMapStatusCol As Long
    heatMapStatusCol = FindColumnInRow(wsHeatMap, 1, Array("Status", "Current Status", "Current Status P1"))
    
    If heatMapStatusCol > 0 Then
        debugReport = debugReport & "✓ Found status column at position " & heatMapStatusCol & vbCrLf
    Else
        debugReport = debugReport & "✗ Could not find status column! Will use column C as default." & vbCrLf
        heatMapStatusCol = 3 ' Default to column C
    End If
    debugReport = debugReport & vbCrLf
    
    ' Step 6: Count operation codes in HeatMap
    debugReport = debugReport & "STEP 6: Scanning operation codes in HeatMap Sheet..." & vbCrLf
    
    Dim heatMapStartRow As Long
    heatMapStartRow = 2 ' Assuming row 1 is header
    
    ' Check if row 1 contains an op code (no header)
    If IsNumeric(wsHeatMap.Cells(1, 1).Value) Then
        heatMapStartRow = 1
        debugReport = debugReport & "No header row detected, starting from row 1" & vbCrLf
    End If
    
    For i = heatMapStartRow To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 8 Then
            heatMapOpCodes = heatMapOpCodes + 1
            If heatMapOpCodes <= 5 Then
                debugReport = debugReport & "  Sample " & heatMapOpCodes & ": Op Code=" & opCode & vbCrLf
            End If
        End If
    Next i
    debugReport = debugReport & "Total operation codes found: " & heatMapOpCodes & vbCrLf & vbCrLf
    
    ' Step 7: Try to match and update
    debugReport = debugReport & "STEP 7: Attempting to match and update statuses..." & vbCrLf
    
    If overallRow > 0 And statusCol > 0 And heatMapOpCodes > 0 Then
        Application.StatusBar = "Updating HeatMap statuses..."
        
        For i = dataStartRow To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Stop at next section
            If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 8 Then
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusCol).Value)))
                
                If finalStatus <> "" And finalStatus <> "N/A" And finalStatus <> "FINAL STATUS" Then
                    ' Search for this opCode in HeatMap
                    For j = heatMapStartRow To lastRowHeatMap
                        If Trim(CStr(wsHeatMap.Cells(j, 1).Value)) = opCode Then
                            matchedCodes = matchedCodes + 1
                            ' Update the status
                            Call SetColoredDot(wsHeatMap.Cells(j, heatMapStatusCol), finalStatus)
                            updatedCount = updatedCount + 1
                            
                            If updatedCount <= 3 Then
                                debugReport = debugReport & "  Updated: " & opCode & " → " & finalStatus & " (row " & j & ")" & vbCrLf
                            End If
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i
        
        debugReport = debugReport & "Matched codes: " & matchedCodes & vbCrLf
        debugReport = debugReport & "Successfully updated: " & updatedCount & vbCrLf & vbCrLf
    Else
        debugReport = debugReport & "Cannot proceed with update due to missing data:" & vbCrLf
        If overallRow = 0 Then debugReport = debugReport & "  - Overall Status section not found" & vbCrLf
        If statusCol = 0 Then debugReport = debugReport & "  - Status column not found" & vbCrLf
        If heatMapOpCodes = 0 Then debugReport = debugReport & "  - No operation codes in HeatMap" & vbCrLf
    End If
    
    ' Cleanup
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Step 8: Show results
    debugReport = debugReport & vbCrLf & "=== SUMMARY ===" & vbCrLf
    debugReport = debugReport & "Time elapsed: " & Format(Timer - startTime, "0.0") & " seconds" & vbCrLf
    debugReport = debugReport & "Operations updated: " & updatedCount & vbCrLf
    
    If updatedCount > 0 Then
        MsgBox debugReport & vbCrLf & vbCrLf & _
               "✓ Successfully updated " & updatedCount & " operation statuses!", _
               vbInformation, "HeatMap Update Complete"
    Else
        MsgBox debugReport & vbCrLf & vbCrLf & _
               "⚠ No operations were updated. Please review the diagnostic information above.", _
               vbExclamation, "HeatMap Update - No Updates"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error in UpdateHeatMapStatus: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl, vbCritical, "Error"
End Sub

' Helper function: Find row containing text
Private Function FindRowContaining(ws As Worksheet, searchText As String, lastRow As Long) As Long
    Dim i As Long
    FindRowContaining = 0
    
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, 1).Value), searchText, vbTextCompare) > 0 Then
            FindRowContaining = i
            Exit Function
        End If
    Next i
End Function

' Helper function: Find column by searching for header text in a specific row
Private Function FindColumnInRow(ws As Worksheet, rowNum As Long, headers As Variant) As Long
    Dim col As Long
    Dim h As Variant
    
    FindColumnInRow = 0
    
    ' Search columns A through Z
    For col = 1 To 26
        For Each h In headers
            If InStr(1, CStr(ws.Cells(rowNum, col).Value), CStr(h), vbTextCompare) > 0 Then
                FindColumnInRow = col
                Exit Function
            End If
        Next h
    Next col
End Function

' Helper function: Set colored dot based on status
Private Sub SetColoredDot(cell As Range, status As String)
    Dim dotChar As String
    Dim dotColor As Long
    
    ' Use filled circle character from Wingdings
    dotChar = "●"
    
    Select Case UCase(Trim(status))
        Case "RED"
            dotColor = RGB(255, 0, 0)     ' Red
        Case "YELLOW"
            dotColor = RGB(255, 192, 0)   ' Yellow/Orange
        Case "GREEN"
            dotColor = RGB(0, 176, 80)    ' Green
        Case Else
            dotColor = RGB(128, 128, 128) ' Gray for N/A or unknown
    End Select
    
    With cell
        .Value = dotChar
        .Font.Name = "Wingdings"
        .Font.Size = 14
        .Font.Color = dotColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' Helper function: Get list of all sheet names
Private Function GetSheetNames() As String
    Dim ws As Worksheet
    Dim names As String
    
    names = ""
    For Each ws In ThisWorkbook.Worksheets
        If names <> "" Then names = names & ", "
        names = names & "'" & ws.Name & "'"
    Next ws
    
    GetSheetNames = names
End Function

' Button creation function
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnExists As Boolean
    
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
        MsgBox "Cannot find HeatMap sheet to add button!", vbExclamation
        Exit Sub
    End If
    
    ' Check if button already exists
    On Error Resume Next
    btnExists = Not ws.Buttons("btnUpdateHeatMap") Is Nothing
    On Error GoTo 0
    
    If btnExists Then
        MsgBox "Button already exists on '" & ws.Name & "' sheet!", vbInformation
        Exit Sub
    End If
    
    ' Create button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    With btn
        .Name = "btnUpdateHeatMap"
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
    End With
    
    MsgBox "Button created successfully on '" & ws.Name & "' sheet!" & vbCrLf & vbCrLf & _
           "Click the button after running evaluation to transfer statuses.", _
           vbInformation, "Button Created"
End Sub
