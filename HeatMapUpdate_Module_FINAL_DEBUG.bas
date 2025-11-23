Attribute VB_Name = "HeatMapUpdate_Module_FINAL_DEBUG"
' ==========================================
' HeatMap Status Update Module - FINAL DEBUG VERSION
' ==========================================
' This module transfers evaluation results from "Evaluation Results" sheet
' to "HeatMap Sheet" with enhanced debugging to diagnose issues
'
' FEATURES:
' - Reads from both "Overall Status by Op Code" and "Operation Mode Summary" sections
' - Updates HeatMap Sheet Status column with colored dots
' - Comprehensive debug messages showing exactly what's happening
' - Handles both sub-operations and parent operations
'
' HOW TO USE:
' 1. Import this module (Alt+F11 → File → Import File)
' 2. Run UpdateHeatMapStatusDebug() macro
' 3. Check the debug messages to see what's being processed
' ==========================================

Option Explicit

' Main function to update HeatMap Status with detailed debugging
Sub UpdateHeatMapStatusDebug()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim evalLastRow As Long, heatLastRow As Long
    Dim i As Long, j As Long
    Dim opCode As String, status As String
    Dim heatOpCode As String
    Dim statusCol As Long
    Dim updateCount As Long
    Dim startTime As Double
    Dim debugMsg As String
    Dim overallSection As Long, summarySection As Long
    Dim evalCol As Long
    Dim foundSheets As Boolean
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updateCount = 0
    debugMsg = "=== HeatMap Status Update - DEBUG MODE ===" & vbCrLf & vbCrLf
    
    ' Step 1: Check if sheets exist
    foundSheets = False
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        debugMsg = debugMsg & "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf
        debugMsg = debugMsg & "Available sheets: "
        For i = 1 To ThisWorkbook.Sheets.Count
            debugMsg = debugMsg & ThisWorkbook.Sheets(i).Name & ", "
        Next i
        MsgBox debugMsg, vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    If wsHeatMap Is Nothing Then
        debugMsg = debugMsg & "ERROR: 'HeatMap Sheet' not found!" & vbCrLf
        debugMsg = debugMsg & "Available sheets: "
        For i = 1 To ThisWorkbook.Sheets.Count
            debugMsg = debugMsg & ThisWorkbook.Sheets(i).Name & ", "
        Next i
        MsgBox debugMsg, vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & "✓ Found both sheets" & vbCrLf & vbCrLf
    
    ' Step 2: Find sections in Evaluation Results
    evalLastRow = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "Evaluation Results last row: " & evalLastRow & vbCrLf
    
    overallSection = 0
    summarySection = 0
    
    ' Search for section headers
    For i = 1 To evalLastRow
        Dim cellValue As String
        cellValue = Trim(wsEval.Cells(i, 1).Value)
        
        If InStr(1, cellValue, "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallSection = i
            debugMsg = debugMsg & "✓ Found 'Overall Status by Op Code' at row " & i & vbCrLf
        ElseIf InStr(1, cellValue, "Operation Mode Summary", vbTextCompare) > 0 Then
            summarySection = i
            debugMsg = debugMsg & "✓ Found 'Operation Mode Summary' at row " & i & vbCrLf
        End If
    Next i
    
    If overallSection = 0 And summarySection = 0 Then
        debugMsg = debugMsg & vbCrLf & "ERROR: Could not find either section!" & vbCrLf
        debugMsg = debugMsg & "Looking for rows containing:" & vbCrLf
        debugMsg = debugMsg & "  - 'Overall Status by Op Code'" & vbCrLf
        debugMsg = debugMsg & "  - 'Operation Mode Summary'" & vbCrLf & vbCrLf
        debugMsg = debugMsg & "First 20 rows in column A:" & vbCrLf
        For i = 1 To Application.Min(20, evalLastRow)
            debugMsg = debugMsg & "Row " & i & ": '" & wsEval.Cells(i, 1).Value & "'" & vbCrLf
        Next i
        MsgBox debugMsg, vbExclamation, "Sections Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' Step 3: Find HeatMap structure
    heatLastRow = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "HeatMap Sheet last row: " & heatLastRow & vbCrLf
    
    ' Find Status column in HeatMap
    statusCol = 0
    For j = 1 To 20 ' Check first 20 columns
        If InStr(1, Trim(wsHeatMap.Cells(1, j).Value), "Status", vbTextCompare) > 0 Then
            statusCol = j
            debugMsg = debugMsg & "✓ Found 'Status' column at column " & j & " (" & ColumnLetter(j) & ")" & vbCrLf
            Exit For
        End If
    Next j
    
    If statusCol = 0 Then
        ' Try column C as default
        statusCol = 3
        debugMsg = debugMsg & "⚠ Status column not found in headers, using column C as default" & vbCrLf
    End If
    
    debugMsg = debugMsg & vbCrLf & "=== Starting Update Process ===" & vbCrLf & vbCrLf
    
    ' Step 4: Process Overall Status by Op Code section
    If overallSection > 0 Then
        debugMsg = debugMsg & "Processing 'Overall Status by Op Code' section..." & vbCrLf
        
        ' Find column with "Overall Status" header
        evalCol = 0
        For j = 1 To 20
            If InStr(1, Trim(wsEval.Cells(overallSection + 1, j).Value), "Overall Status", vbTextCompare) > 0 Then
                evalCol = j
                debugMsg = debugMsg & "  Status column: " & ColumnLetter(j) & vbCrLf
                Exit For
            End If
        Next j
        
        If evalCol > 0 Then
            ' Process data rows (skip header row)
            For i = overallSection + 2 To evalLastRow
                ' Stop if we hit the next section
                If Trim(wsEval.Cells(i, 1).Value) = "Operation Mode Summary" Or _
                   Trim(wsEval.Cells(i, 1).Value) = "" Then
                    Exit For
                End If
                
                opCode = Trim(wsEval.Cells(i, 1).Value)
                status = Trim(wsEval.Cells(i, evalCol).Value)
                
                ' Skip if empty
                If opCode <> "" And status <> "" Then
                    ' Find matching row in HeatMap
                    For j = 1 To heatLastRow
                        heatOpCode = Trim(wsHeatMap.Cells(j, 1).Value)
                        If heatOpCode = opCode Then
                            ' Update status
                            Call SetStatusDot(wsHeatMap, j, statusCol, status)
                            updateCount = updateCount + 1
                            debugMsg = debugMsg & "  ✓ Updated " & opCode & " → " & status & vbCrLf
                            Exit For
                        End If
                    Next j
                End If
            Next i
        Else
            debugMsg = debugMsg & "  ⚠ Could not find 'Overall Status' column" & vbCrLf
        End If
        
        debugMsg = debugMsg & vbCrLf
    End If
    
    ' Step 5: Process Operation Mode Summary section
    If summarySection > 0 Then
        debugMsg = debugMsg & "Processing 'Operation Mode Summary' section..." & vbCrLf
        
        ' Find column with "Final Status" header
        evalCol = 0
        For j = 1 To 20
            If InStr(1, Trim(wsEval.Cells(summarySection + 1, j).Value), "Final Status", vbTextCompare) > 0 Then
                evalCol = j
                debugMsg = debugMsg & "  Status column: " & ColumnLetter(j) & vbCrLf
                Exit For
            End If
        Next j
        
        If evalCol > 0 Then
            ' Process data rows (skip header row)
            For i = summarySection + 2 To evalLastRow
                ' Stop if empty row
                If Trim(wsEval.Cells(i, 1).Value) = "" Then
                    Exit For
                End If
                
                opCode = Trim(wsEval.Cells(i, 1).Value)
                status = Trim(wsEval.Cells(i, evalCol).Value)
                
                ' Skip if empty
                If opCode <> "" And status <> "" Then
                    ' Find matching row in HeatMap
                    For j = 1 To heatLastRow
                        heatOpCode = Trim(wsHeatMap.Cells(j, 1).Value)
                        If heatOpCode = opCode Then
                            ' Update status
                            Call SetStatusDot(wsHeatMap, j, statusCol, status)
                            updateCount = updateCount + 1
                            debugMsg = debugMsg & "  ✓ Updated " & opCode & " → " & status & vbCrLf
                            Exit For
                        End If
                    Next j
                End If
            Next i
        Else
            debugMsg = debugMsg & "  ⚠ Could not find 'Final Status' column" & vbCrLf
        End If
    End If
    
    ' Step 6: Summary
    debugMsg = debugMsg & vbCrLf & "=== Update Complete ===" & vbCrLf
    debugMsg = debugMsg & "Total operations updated: " & updateCount & vbCrLf
    debugMsg = debugMsg & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf
    
    MsgBox debugMsg, vbInformation, "HeatMap Status Update - Debug Results"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error: " & Err.Description & vbCrLf & vbCrLf & debugMsg, vbCritical, "Update Error"
End Sub

' Helper function to set colored status dot
Private Sub SetStatusDot(ws As Worksheet, row As Long, col As Long, status As String)
    Dim dotChar As String
    Dim dotColor As Long
    
    status = UCase(Trim(status))
    dotChar = ChrW(9679) ' Filled circle character ●
    
    ' Determine color based on status
    Select Case status
        Case "RED"
            dotColor = RGB(255, 0, 0) ' Red
        Case "YELLOW"
            dotColor = RGB(255, 192, 0) ' Yellow
        Case "GREEN"
            dotColor = RGB(0, 176, 80) ' Green
        Case "N/A", ""
            dotColor = RGB(128, 128, 128) ' Gray
        Case Else
            dotColor = RGB(128, 128, 128) ' Gray for unknown
    End Select
    
    ' Set the cell value and formatting
    With ws.Cells(row, col)
        .Value = dotChar
        .Font.Name = "Wingdings"
        .Font.Size = 14
        .Font.Color = dotColor
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
End Sub

' Helper function to convert column number to letter
Private Function ColumnLetter(colNum As Long) As String
    Dim temp As Long
    Dim letter As String
    
    temp = colNum
    Do While temp > 0
        temp = temp - 1
        letter = Chr((temp Mod 26) + 65) & letter
        temp = temp \ 26
    Loop
    
    ColumnLetter = letter
End Function

' Function to create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim wsHeatMap As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If wsHeatMap Is Nothing Then
        MsgBox "'HeatMap Sheet' not found!", vbExclamation
        Exit Sub
    End If
    
    ' Remove existing button if present
    btnName = "btnUpdateHeatMap"
    On Error Resume Next
    wsHeatMap.Buttons(btnName).Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = wsHeatMap.Buttons.Add(10, 10, 200, 30)
    btn.Name = btnName
    btn.Text = "Update HeatMap Status (DEBUG)"
    btn.OnAction = "UpdateHeatMapStatusDebug"
    
    MsgBox "Button created on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update status with detailed debug information.", _
           vbInformation, "Button Created"
End Sub
