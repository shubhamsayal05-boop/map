Attribute VB_Name = "HeatMapUpdate_Enhanced"
' ====================================================================
' Module: HeatMapUpdate_Enhanced
' Purpose: Transfer evaluation results to HeatMap Sheet with detailed debugging
' Version: Enhanced Final with proper section detection
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
    Dim statusColHeatMap As Long
    Dim opCodeColEval As Long
    Dim statusColEval As Long
    Dim summaryStartRow As Long
    Dim overallStatusStartRow As Long
    Dim foundInEval As Boolean
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugMsg = ""
    
    ' Step 1: Get worksheets
    debugMsg = debugMsg & "Step 1: Looking for sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetNames(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & "  ✓ Found 'Evaluation Results' sheet" & vbCrLf
    debugMsg = debugMsg & "  ✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    
    ' Step 2: Find sections in Evaluation Results sheet
    debugMsg = debugMsg & "Step 2: Finding sections in Evaluation Results..." & vbCrLf
    
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, 1).End(xlUp).Row
    overallStatusStartRow = 0
    summaryStartRow = 0
    
    ' Look for "Overall Status by Op Code" section
    For i = 1 To lastRowEval
        If InStr(1, wsEval.Cells(i, 1).Value, "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallStatusStartRow = i + 1 ' Data starts next row
            debugMsg = debugMsg & "  ✓ Found 'Overall Status by Op Code' at row " & i & vbCrLf
            Exit For
        End If
    Next i
    
    ' Look for "Operation Mode Summary" section
    For i = 1 To lastRowEval
        If InStr(1, wsEval.Cells(i, 1).Value, "Operation Mode Summary", vbTextCompare) > 0 Then
            summaryStartRow = i + 1 ' Data starts next row
            debugMsg = debugMsg & "  ✓ Found 'Operation Mode Summary' at row " & i & vbCrLf
            Exit For
        End If
    Next i
    
    If overallStatusStartRow = 0 And summaryStartRow = 0 Then
        MsgBox "ERROR: Could not find sections!" & vbCrLf & vbCrLf & _
               "Looking for:" & vbCrLf & _
               "  - 'Overall Status by Op Code'" & vbCrLf & _
               "  - 'Operation Mode Summary'" & vbCrLf & vbCrLf & _
               debugMsg, vbCritical, "Sections Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' Step 3: Find Op Code and Status columns in Evaluation Results
    debugMsg = debugMsg & "Step 3: Finding columns in Evaluation Results..." & vbCrLf
    
    ' Use whichever section we found first
    Dim headerRow As Long
    If overallStatusStartRow > 0 Then
        headerRow = overallStatusStartRow - 1
    Else
        headerRow = summaryStartRow - 1
    End If
    
    opCodeColEval = 0
    statusColEval = 0
    
    For i = 1 To 20 ' Check first 20 columns
        Dim headerVal As String
        headerVal = Trim(UCase(wsEval.Cells(headerRow, i).Value))
        
        If headerVal = "OP CODE" Or headerVal = "OPCODE" Or headerVal = "CODE" Then
            opCodeColEval = i
            debugMsg = debugMsg & "  ✓ Found Op Code column at column " & i & vbCrLf
        End If
        
        If InStr(1, headerVal, "OVERALL STATUS", vbTextCompare) > 0 Or _
           InStr(1, headerVal, "FINAL STATUS", vbTextCompare) > 0 Or _
           headerVal = "STATUS" Then
            statusColEval = i
            debugMsg = debugMsg & "  ✓ Found Status column at column " & i & vbCrLf
        End If
    Next i
    
    If opCodeColEval = 0 Then
        MsgBox "ERROR: Could not find 'Op Code' column in Evaluation Results!" & vbCrLf & vbCrLf & debugMsg, _
               vbCritical, "Column Not Found"
        Exit Sub
    End If
    
    If statusColEval = 0 Then
        MsgBox "ERROR: Could not find Status column in Evaluation Results!" & vbCrLf & vbCrLf & debugMsg, _
               vbCritical, "Column Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' Step 4: Find Status column in HeatMap Sheet
    debugMsg = debugMsg & "Step 4: Finding Status column in HeatMap Sheet..." & vbCrLf
    
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, 1).End(xlUp).Row
    statusColHeatMap = 0
    
    ' Look in first row for "Status" or "Current Status" header
    For i = 1 To 20
        headerVal = Trim(UCase(wsHeatMap.Cells(1, i).Value))
        If InStr(1, headerVal, "STATUS", vbTextCompare) > 0 Then
            statusColHeatMap = i
            debugMsg = debugMsg & "  ✓ Found Status column at column " & i & " ('" & wsHeatMap.Cells(1, i).Value & "')" & vbCrLf
            Exit For
        End If
    Next i
    
    If statusColHeatMap = 0 Then
        MsgBox "ERROR: Could not find 'Status' column in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Please ensure HeatMap Sheet has a column with 'Status' in the header." & vbCrLf & vbCrLf & _
               debugMsg, vbCritical, "Column Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' Step 5: Update HeatMap statuses
    debugMsg = debugMsg & "Step 5: Updating statuses..." & vbCrLf
    
    ' Process each row in HeatMap Sheet
    For i = 2 To lastRowHeatMap ' Start from row 2 (skip header)
        opCode = Trim(wsHeatMap.Cells(i, 1).Value) ' Column A has Op Codes
        
        If opCode <> "" And IsNumeric(opCode) Then
            foundInEval = False
            finalStatus = ""
            
            ' Search in Overall Status section
            If overallStatusStartRow > 0 Then
                For j = overallStatusStartRow To lastRowEval
                    If Trim(wsEval.Cells(j, opCodeColEval).Value) = opCode Then
                        finalStatus = Trim(wsEval.Cells(j, statusColEval).Value)
                        foundInEval = True
                        Exit For
                    End If
                    
                    ' Stop if we hit next section
                    If InStr(1, wsEval.Cells(j, 1).Value, "Operation Mode Summary", vbTextCompare) > 0 Then
                        Exit For
                    End If
                Next j
            End If
            
            ' If not found in Overall Status, search in Summary section
            If Not foundInEval And summaryStartRow > 0 Then
                For j = summaryStartRow To lastRowEval
                    If Trim(wsEval.Cells(j, opCodeColEval).Value) = opCode Then
                        finalStatus = Trim(wsEval.Cells(j, statusColEval).Value)
                        foundInEval = True
                        Exit For
                    End If
                    
                    ' Stop if we hit empty rows
                    If Trim(wsEval.Cells(j, 1).Value) = "" And Trim(wsEval.Cells(j, 2).Value) = "" Then
                        Exit For
                    End If
                Next j
            End If
            
            ' Update if found
            If foundInEval And finalStatus <> "" Then
                ' Set colored dot based on status
                Call SetStatusDot(wsHeatMap.Cells(i, statusColHeatMap), finalStatus)
                updatedCount = updatedCount + 1
            End If
        End If
    Next i
    
    debugMsg = debugMsg & "  ✓ Updated " & updatedCount & " operations" & vbCrLf
    
    ' Show results
    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 2)
    
    MsgBox "HeatMap Status Update Complete!" & vbCrLf & vbCrLf & _
           "Updated: " & updatedCount & " operations" & vbCrLf & _
           "Time: " & elapsed & " seconds" & vbCrLf & vbCrLf & _
           "Details:" & vbCrLf & debugMsg, _
           vbInformation, "Success"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error updating HeatMap!" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description & vbCrLf & _
           "Number: " & Err.Number & vbCrLf & vbCrLf & _
           "Debug info:" & vbCrLf & debugMsg, _
           vbCritical, "Error"
End Sub

' Helper function to set colored status dot
Private Sub SetStatusDot(cell As Range, status As String)
    Dim statusUpper As String
    statusUpper = UCase(Trim(status))
    
    With cell
        .Font.Name = "Wingdings"
        .Font.Size = 14
        .Value = "l" ' Filled circle character in Wingdings
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        
        ' Set color based on status
        Select Case statusUpper
            Case "RED"
                .Font.Color = RGB(255, 0, 0) ' RED
            Case "YELLOW"
                .Font.Color = RGB(255, 192, 0) ' YELLOW/ORANGE
            Case "GREEN"
                .Font.Color = RGB(0, 176, 80) ' GREEN
            Case Else ' N/A or blank
                .Font.Color = RGB(128, 128, 128) ' GRAY
        End Select
    End With
End Sub

' Helper function to get list of sheet names
Private Function GetSheetNames() As String
    Dim ws As Worksheet
    Dim names As String
    names = ""
    
    For Each ws In ThisWorkbook.Sheets
        If names <> "" Then names = names & ", "
        names = names & "'" & ws.Name & "'"
    Next ws
    
    GetSheetNames = names
End Function

' Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found! Please create it first.", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing button if present
    btnName = "btnUpdateHeatMap"
    On Error Resume Next
    ws.Buttons(btnName).Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    With btn
        .Name = btnName
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button after running evaluation to fill in statuses.", _
           vbInformation, "Button Created"
End Sub
