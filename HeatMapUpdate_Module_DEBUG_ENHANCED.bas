Attribute VB_Name = "HeatMapUpdate_Debug"
' ====================================================================
' Module: HeatMapUpdate_Debug
' Purpose: Transfer evaluation results to HeatMap Sheet with ENHANCED DEBUGGING
' ====================================================================

Option Explicit

' Main function with comprehensive debugging
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
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugInfo = "=== HEATMAP UPDATE DEBUG INFO ===" & vbCrLf & vbCrLf
    
    ' Step 1: Find sheets
    debugInfo = debugInfo & "STEP 1: Finding Sheets" & vbCrLf
    debugInfo = debugInfo & String(50, "-") & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        ' Try alternate names
        Set wsEval = ThisWorkbook.Sheets("Evaluation Result")
        If wsEval Is Nothing Then
            debugInfo = debugInfo & "ERROR: Cannot find 'Evaluation Results' sheet!" & vbCrLf
            debugInfo = debugInfo & "Available sheets:" & vbCrLf
            Dim ws As Worksheet
            For Each ws In ThisWorkbook.Worksheets
                debugInfo = debugInfo & "  - " & ws.Name & vbCrLf
            Next ws
            MsgBox debugInfo, vbCritical, "Sheet Not Found"
            Exit Sub
        End If
    End If
    debugInfo = debugInfo & "✓ Found Evaluation sheet: " & wsEval.Name & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
        If wsHeatMap Is Nothing Then
            debugInfo = debugInfo & "ERROR: Cannot find 'HeatMap Sheet'!" & vbCrLf
            MsgBox debugInfo, vbCritical, "Sheet Not Found"
            Exit Sub
        End If
    End If
    debugInfo = debugInfo & "✓ Found HeatMap sheet: " & wsHeatMap.Name & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    ' Step 2: Analyze Evaluation Results structure
    debugInfo = debugInfo & "STEP 2: Analyzing Evaluation Results Sheet" & vbCrLf
    debugInfo = debugInfo & String(50, "-") & vbCrLf
    
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "Last row in Evaluation Results: " & lastRowEval & vbCrLf
    
    ' Show first 10 rows of column A
    debugInfo = debugInfo & vbCrLf & "First 10 rows in Column A:" & vbCrLf
    For i = 1 To Application.Min(10, lastRowEval)
        debugInfo = debugInfo & "  Row " & i & ": [" & wsEval.Cells(i, 1).Value & "]" & vbCrLf
    Next i
    
    ' Find "Overall Status by Op Code" section
    Dim overallRow As Long, summaryRow As Long
    overallRow = 0
    summaryRow = 0
    
    For i = 1 To lastRowEval
        Dim cellValue As String
        cellValue = Trim(CStr(wsEval.Cells(i, 1).Value))
        If InStr(1, cellValue, "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallRow = i
            debugInfo = debugInfo & vbCrLf & "✓ Found 'Overall Status by Op Code' at row: " & i & vbCrLf
        End If
        If InStr(1, cellValue, "Operation Mode Summary", vbTextCompare) > 0 Then
            summaryRow = i
            debugInfo = debugInfo & "✓ Found 'Operation Mode Summary' at row: " & i & vbCrLf
        End If
        If overallRow > 0 And summaryRow > 0 Then Exit For
    Next i
    
    If overallRow = 0 Then
        debugInfo = debugInfo & vbCrLf & "ERROR: Cannot find 'Overall Status by Op Code' section!" & vbCrLf
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Step 3: Find status column header
    debugInfo = debugInfo & "STEP 3: Finding Status Column" & vbCrLf
    debugInfo = debugInfo & String(50, "-") & vbCrLf
    
    If overallRow > 0 Then
        Dim headerRow As Long
        headerRow = overallRow + 1  ' Header is usually next row
        
        debugInfo = debugInfo & "Checking row " & headerRow & " for headers:" & vbCrLf
        For j = 1 To 20  ' Check first 20 columns
            Dim headerVal As String
            headerVal = Trim(CStr(wsEval.Cells(headerRow, j).Value))
            If headerVal <> "" Then
                debugInfo = debugInfo & "  Col " & j & " (" & Split(Cells(1, j).Address, "$")(1) & "): [" & headerVal & "]" & vbCrLf
                If InStr(1, headerVal, "Final Status", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Overall Status", vbTextCompare) > 0 Or _
                   InStr(1, headerVal, "Status", vbTextCompare) > 0 Then
                    statusCol = j
                End If
            End If
        Next j
        
        If statusCol > 0 Then
            debugInfo = debugInfo & vbCrLf & "✓ Using status column: " & statusCol & vbCrLf
        Else
            debugInfo = debugInfo & vbCrLf & "WARNING: Could not find status column. Trying column C (3)..." & vbCrLf
            statusCol = 3  ' Default fallback
        End If
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Step 4: Analyze HeatMap structure
    debugInfo = debugInfo & "STEP 4: Analyzing HeatMap Sheet" & vbCrLf
    debugInfo = debugInfo & String(50, "-") & vbCrLf
    
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "Last row in HeatMap: " & lastRowHeatMap & vbCrLf
    
    ' Show first 10 rows
    debugInfo = debugInfo & vbCrLf & "First 10 rows in HeatMap Column A:" & vbCrLf
    For i = 1 To Application.Min(10, lastRowHeatMap)
        debugInfo = debugInfo & "  Row " & i & ": [" & wsHeatMap.Cells(i, 1).Value & "]" & vbCrLf
    Next i
    
    ' Find status column in HeatMap
    Dim heatMapStatusCol As Long
    heatMapStatusCol = 0
    debugInfo = debugInfo & vbCrLf & "Looking for status column in HeatMap (Row 1):" & vbCrLf
    For j = 1 To 20
        Dim heatMapHeader As String
        heatMapHeader = Trim(CStr(wsHeatMap.Cells(1, j).Value))
        If heatMapHeader <> "" Then
            debugInfo = debugInfo & "  Col " & j & ": [" & heatMapHeader & "]" & vbCrLf
            If InStr(1, heatMapHeader, "Status", vbTextCompare) > 0 And _
               InStr(1, heatMapHeader, "Current", vbTextCompare) > 0 Then
                heatMapStatusCol = j
            End If
        End If
    Next j
    
    If heatMapStatusCol = 0 Then
        ' Try row 2 or 3 for headers
        For i = 2 To 3
            For j = 1 To 20
                heatMapHeader = Trim(CStr(wsHeatMap.Cells(i, j).Value))
                If InStr(1, heatMapHeader, "Status", vbTextCompare) > 0 Then
                    heatMapStatusCol = j
                    debugInfo = debugInfo & vbCrLf & "Found status column in row " & i & ", col " & j & vbCrLf
                    Exit For
                End If
            Next j
            If heatMapStatusCol > 0 Then Exit For
        Next i
    End If
    
    If heatMapStatusCol = 0 Then
        debugInfo = debugInfo & vbCrLf & "WARNING: Could not find status column. Using column C (3)..." & vbCrLf
        heatMapStatusCol = 3  ' Default
    Else
        debugInfo = debugInfo & vbCrLf & "✓ Using HeatMap status column: " & heatMapStatusCol & vbCrLf
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Step 5: Process data
    debugInfo = debugInfo & "STEP 5: Processing Operations" & vbCrLf
    debugInfo = debugInfo & String(50, "-") & vbCrLf
    
    If overallRow > 0 And statusCol > 0 Then
        Dim dataStartRow As Long
        dataStartRow = overallRow + 2  ' Skip section title and header
        
        debugInfo = debugInfo & "Processing from row " & dataStartRow & " in Evaluation Results" & vbCrLf
        debugInfo = debugInfo & vbCrLf & "Sample data:" & vbCrLf
        
        Dim sampleCount As Long
        sampleCount = 0
        
        For i = dataStartRow To lastRowEval
            ' Stop at next section
            If summaryRow > 0 And i >= summaryRow Then
                debugInfo = debugInfo & "Reached Operation Mode Summary section" & vbCrLf
                Exit For
            End If
            
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Skip empty or non-numeric
            If opCode = "" Or Not IsNumeric(opCode) Then
                Continue For
            End If
            
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusCol).Value)))
            
            ' Show first 3 samples
            If sampleCount < 3 Then
                debugInfo = debugInfo & "  Row " & i & ": OpCode=" & opCode & ", Status=" & finalStatus & vbCrLf
                sampleCount = sampleCount + 1
            End If
            
            ' Skip if no valid status
            If finalStatus = "" Or finalStatus = "N/A" Or finalStatus = "FINAL STATUS" Then
                Continue For
            End If
            
            ' Find and update in HeatMap
            For j = 1 To lastRowHeatMap
                Dim heatMapCode As String
                heatMapCode = Trim(CStr(wsHeatMap.Cells(j, 1).Value))
                
                If heatMapCode = opCode Then
                    ' Update status with colored dot
                    wsHeatMap.Cells(j, heatMapStatusCol).Value = GetStatusDot(finalStatus)
                    wsHeatMap.Cells(j, heatMapStatusCol).Font.Name = "Wingdings"
                    wsHeatMap.Cells(j, heatMapStatusCol).Font.Size = 14
                    wsHeatMap.Cells(j, heatMapStatusCol).Font.Color = GetStatusColor(finalStatus)
                    updatedCount = updatedCount + 1
                    Exit For
                End If
            Next j
        Next i
    End If
    
    debugInfo = debugInfo & vbCrLf & "Processing complete!" & vbCrLf
    debugInfo = debugInfo & "Operations updated: " & updatedCount & vbCrLf
    debugInfo = debugInfo & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf
    
    ' Show results
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    If updatedCount = 0 Then
        MsgBox debugInfo, vbExclamation, "Debug Information - No Updates"
    Else
        MsgBox debugInfo, vbInformation, "Debug Information - Success!"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & debugInfo, vbCritical, "Error"
End Sub

' Helper function to get colored dot based on status
Function GetStatusDot(status As String) As String
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusDot = "l"  ' Filled circle in Wingdings
        Case "YELLOW"
            GetStatusDot = "l"
        Case "GREEN"
            GetStatusDot = "l"
        Case Else
            GetStatusDot = "l"  ' Gray for N/A or others
    End Select
End Function

' Helper function to get color code for status
Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0)       ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 192, 0)     ' Yellow/Orange
        Case "GREEN"
            GetStatusColor = RGB(0, 176, 80)      ' Green
        Case Else
            GetStatusColor = RGB(128, 128, 128)   ' Gray for N/A
    End Select
End Function

' Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim wsHeatMap As Worksheet
    Dim btn As Button
    
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
    End If
    On Error GoTo 0
    
    If wsHeatMap Is Nothing Then
        MsgBox "Cannot find HeatMap Sheet!", vbCritical
        Exit Sub
    End If
    
    ' Remove existing button if present
    On Error Resume Next
    wsHeatMap.Buttons("UpdateHeatMapButton").Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = wsHeatMap.Buttons.Add(10, 10, 150, 30)
    btn.Name = "UpdateHeatMapButton"
    btn.Caption = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on " & wsHeatMap.Name & "!" & vbCrLf & vbCrLf & _
           "Click the button to update status with debug information.", _
           vbInformation, "Button Created"
End Sub
