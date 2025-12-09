Attribute VB_Name = "HeatMapUpdate_Enhanced"
' ====================================================================
' Module: HeatMapUpdate_Enhanced
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Version: Enhanced with detailed diagnostics
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
    Dim debugInfo As String
    Dim statusCol As Long
    Dim opCodeCol As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugInfo = "=== HEATMAP UPDATE DIAGNOSTIC ===" & vbCrLf & vbCrLf
    
    ' Step 1: Verify sheets exist
    Application.StatusBar = "Step 1: Verifying sheets..."
    debugInfo = debugInfo & "Step 1: Verifying sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetNames(), vbCritical
        Exit Sub
    End If
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetNames(), vbCritical
        Exit Sub
    End If
    
    debugInfo = debugInfo & "✓ Found Evaluation Results sheet" & vbCrLf
    debugInfo = debugInfo & "✓ Found HeatMap Sheet" & vbCrLf & vbCrLf
    
    ' Step 2: Analyze Evaluation Results structure
    Application.StatusBar = "Step 2: Analyzing Evaluation Results..."
    debugInfo = debugInfo & "Step 2: Analyzing Evaluation Results..." & vbCrLf
    
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "Last row with data: " & lastRowEval & vbCrLf
    
    ' Show first 10 rows of Evaluation Results
    debugInfo = debugInfo & vbCrLf & "First 10 rows of Evaluation Results (Col A-C):" & vbCrLf
    For i = 1 To Application.Min(10, lastRowEval)
        debugInfo = debugInfo & "Row " & i & ": [" & _
                    Trim(CStr(wsEval.Cells(i, 1).Value)) & "] [" & _
                    Trim(CStr(wsEval.Cells(i, 2).Value)) & "] [" & _
                    Trim(CStr(wsEval.Cells(i, 3).Value)) & "]" & vbCrLf
    Next i
    debugInfo = debugInfo & vbCrLf
    
    ' Step 3: Find "Overall Status by Op Code" section
    Application.StatusBar = "Step 3: Finding Overall Status section..."
    debugInfo = debugInfo & "Step 3: Finding 'Overall Status by Op Code' section..." & vbCrLf
    
    Dim overallStartRow As Long
    overallStartRow = 0
    
    For i = 1 To lastRowEval
        If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallStartRow = i
            Exit For
        End If
    Next i
    
    If overallStartRow = 0 Then
        MsgBox "ERROR: Could not find 'Overall Status by Op Code' section!" & vbCrLf & vbCrLf & _
               "Debug Info:" & vbCrLf & debugInfo, vbCritical
        Exit Sub
    End If
    
    debugInfo = debugInfo & "✓ Found at row: " & overallStartRow & vbCrLf
    
    ' Show header row
    Dim headerRow As Long
    headerRow = overallStartRow + 1
    debugInfo = debugInfo & "Header row " & headerRow & ": "
    For i = 1 To 15
        debugInfo = debugInfo & "[" & Trim(CStr(wsEval.Cells(headerRow, i).Value)) & "] "
    Next i
    debugInfo = debugInfo & vbCrLf & vbCrLf
    
    ' Step 4: Find Final Status column
    Application.StatusBar = "Step 4: Finding Final Status column..."
    debugInfo = debugInfo & "Step 4: Finding Final Status column..." & vbCrLf
    
    statusCol = 0
    For i = 1 To 20
        Dim headerValue As String
        headerValue = Trim(UCase(CStr(wsEval.Cells(headerRow, i).Value)))
        If headerValue = "FINAL STATUS" Or headerValue = "OVERALL STATUS" Then
            statusCol = i
            Exit For
        End If
    Next i
    
    If statusCol = 0 Then
        MsgBox "ERROR: Could not find 'Final Status' or 'Overall Status' column!" & vbCrLf & vbCrLf & _
               "Debug Info:" & vbCrLf & debugInfo, vbCritical
        Exit Sub
    End If
    
    debugInfo = debugInfo & "✓ Found at column: " & statusCol & " (" & Chr(64 + statusCol) & ")" & vbCrLf & vbCrLf
    
    ' Step 5: Analyze HeatMap Sheet structure
    Application.StatusBar = "Step 5: Analyzing HeatMap Sheet..."
    debugInfo = debugInfo & "Step 5: Analyzing HeatMap Sheet..." & vbCrLf
    
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "Last row with data: " & lastRowHeatMap & vbCrLf
    
    ' Show first 10 rows
    debugInfo = debugInfo & vbCrLf & "First 10 rows of HeatMap Sheet (Col A-C):" & vbCrLf
    For i = 1 To Application.Min(10, lastRowHeatMap)
        debugInfo = debugInfo & "Row " & i & ": [" & _
                    Trim(CStr(wsHeatMap.Cells(i, 1).Value)) & "] [" & _
                    Trim(CStr(wsHeatMap.Cells(i, 2).Value)) & "] [" & _
                    Trim(CStr(wsHeatMap.Cells(i, 3).Value)) & "]" & vbCrLf
    Next i
    debugInfo = debugInfo & vbCrLf
    
    ' Step 6: Find Status column in HeatMap
    Application.StatusBar = "Step 6: Finding Status column in HeatMap..."
    debugInfo = debugInfo & "Step 6: Finding Status column in HeatMap..." & vbCrLf
    
    Dim heatMapStatusCol As Long
    heatMapStatusCol = 0
    
    ' Look in first few rows for header
    For i = 1 To Application.Min(5, lastRowHeatMap)
        For j = 1 To 20
            Dim cellValue As String
            cellValue = Trim(UCase(CStr(wsHeatMap.Cells(i, j).Value)))
            If cellValue = "STATUS" Or cellValue = "CURRENT STATUS" Then
                heatMapStatusCol = j
                Exit For
            End If
        Next j
        If heatMapStatusCol > 0 Then Exit For
    Next i
    
    If heatMapStatusCol = 0 Then
        MsgBox "ERROR: Could not find 'Status' column in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Debug Info:" & vbCrLf & debugInfo, vbCritical
        Exit Sub
    End If
    
    debugInfo = debugInfo & "✓ Found at column: " & heatMapStatusCol & " (" & Chr(64 + heatMapStatusCol) & ")" & vbCrLf & vbCrLf
    
    ' Step 7: Process and update statuses
    Application.StatusBar = "Step 7: Updating statuses..."
    debugInfo = debugInfo & "Step 7: Processing and updating statuses..." & vbCrLf
    
    Application.ScreenUpdating = False
    
    Dim dataStartRow As Long
    dataStartRow = overallStartRow + 2 ' Skip section title and header
    
    Dim processedOps As String
    processedOps = ""
    
    For i = dataStartRow To lastRowEval
        opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
        
        ' Stop if we hit Operation Mode Summary or another section
        If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Or _
           InStr(1, opCode, "Accelerations", vbTextCompare) > 0 Or _
           InStr(1, opCode, "Decelerations", vbTextCompare) > 0 Then
            Exit For
        End If
        
        ' Process only if it looks like an operation code (numeric)
        If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 7 Then
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusCol).Value)))
            
            ' Skip headers and N/A
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                ' Find matching row in HeatMap
                Dim matchFound As Boolean
                matchFound = False
                
                For j = 1 To lastRowHeatMap
                    Dim heatMapOpCode As String
                    heatMapOpCode = Trim(CStr(wsHeatMap.Cells(j, 1).Value))
                    
                    If heatMapOpCode = opCode Then
                        ' Update status with colored dot
                        wsHeatMap.Cells(j, heatMapStatusCol).Value = GetStatusSymbol(finalStatus)
                        wsHeatMap.Cells(j, heatMapStatusCol).Font.Name = "Wingdings"
                        wsHeatMap.Cells(j, heatMapStatusCol).Font.Size = 14
                        wsHeatMap.Cells(j, heatMapStatusCol).Font.Color = GetStatusColor(finalStatus)
                        
                        updatedCount = updatedCount + 1
                        matchFound = True
                        processedOps = processedOps & opCode & " -> " & finalStatus & vbCrLf
                        Exit For
                    End If
                Next j
                
                If Not matchFound Then
                    processedOps = processedOps & opCode & " -> NOT FOUND IN HEATMAP!" & vbCrLf
                End If
            End If
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Final results
    debugInfo = debugInfo & vbCrLf & "Processing Results:" & vbCrLf
    debugInfo = debugInfo & "Operations updated: " & updatedCount & vbCrLf
    debugInfo = debugInfo & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf & vbCrLf
    debugInfo = debugInfo & "Processed Operations:" & vbCrLf & processedOps
    
    ' Show results
    If updatedCount > 0 Then
        MsgBox "HeatMap Status Update Complete!" & vbCrLf & vbCrLf & _
               "Operations updated: " & updatedCount & vbCrLf & _
               "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf & vbCrLf & _
               "Click OK to see detailed diagnostic info.", vbInformation
        
        ' Show debug info in a separate message
        MsgBox debugInfo, vbInformation, "Diagnostic Information"
    Else
        MsgBox "WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               "Debug Info:" & vbCrLf & vbCrLf & debugInfo, vbExclamation
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "Debug Info:" & vbCrLf & debugInfo, vbCritical
End Sub

' Helper function to get sheet names
Private Function GetSheetNames() As String
    Dim ws As Worksheet
    Dim names As String
    names = ""
    For Each ws In ThisWorkbook.Worksheets
        names = names & ws.Name & ", "
    Next ws
    If Len(names) > 2 Then names = Left(names, Len(names) - 2)
    GetSheetNames = names
End Function

' Helper function to get status symbol
Private Function GetStatusSymbol(status As String) As String
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusSymbol = "l" ' Filled circle in Wingdings
        Case "YELLOW"
            GetStatusSymbol = "l" ' Filled circle in Wingdings
        Case "GREEN"
            GetStatusSymbol = "l" ' Filled circle in Wingdings
        Case Else
            GetStatusSymbol = "l" ' Filled circle in Wingdings (gray)
    End Select
End Function

' Helper function to get status color
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

' Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnExists As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found! Please create it first.", vbExclamation
        Exit Sub
    End If
    
    ' Check if button already exists
    btnExists = False
    On Error Resume Next
    Set btn = ws.Buttons("UpdateHeatMapButton")
    If Not btn Is Nothing Then btnExists = True
    On Error GoTo 0
    
    If btnExists Then
        MsgBox "Button already exists on HeatMap Sheet!", vbInformation
        Exit Sub
    End If
    
    ' Create button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = "UpdateHeatMapButton"
    btn.Text = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    ' Format button
    btn.Font.Bold = True
    btn.Font.Size = 10
    
    MsgBox "Button 'Update HeatMap Status' created on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to transfer evaluation results to HeatMap.", vbInformation
End Sub
