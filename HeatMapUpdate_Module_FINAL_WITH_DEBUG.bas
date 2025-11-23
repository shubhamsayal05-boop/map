Attribute VB_Name = "HeatMapUpdate_Module_FINAL_WITH_DEBUG"
' ============================================================================
' Module: HeatMapUpdate_Module_FINAL_WITH_DEBUG
' Purpose: Transfer evaluation results to HeatMap Sheet with enhanced debugging
' Version: Final with comprehensive diagnostics
' ============================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results
Public Sub UpdateHeatMapStatus()
    On Error GoTo ErrorHandler
    
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim evalLastRow As Long
    Dim heatMapLastRow As Long
    Dim i As Long, j As Long
    Dim opCode As String
    Dim status As String
    Dim updateCount As Long
    Dim startTime As Double
    Dim statusCol As Long
    Dim evalStartRow As Long
    Dim summaryStartRow As Long
    Dim debugMsg As String
    
    startTime = Timer
    updateCount = 0
    
    ' Step 1: Find and activate sheets
    debugMsg = "Step 1: Finding sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Please ensure:" & vbCrLf & _
               "1. Sheet name is exactly 'Evaluation Results'" & vbCrLf & _
               "2. Evaluation has been run", vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Please ensure sheet name is exactly 'HeatMap Sheet'", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    ' Step 2: Find "Overall Status by Op Code" section in Evaluation Results
    debugMsg = debugMsg & "Step 2: Locating 'Overall Status by Op Code' section..." & vbCrLf
    evalStartRow = 0
    evalLastRow = wsEval.Cells(wsEval.Rows.Count, 1).End(xlUp).Row
    
    For i = 1 To evalLastRow
        If InStr(1, wsEval.Cells(i, 1).Value, "Overall Status by Op Code", vbTextCompare) > 0 Then
            evalStartRow = i + 2 ' Data starts 2 rows after header
            debugMsg = debugMsg & "✓ Found at row " & i & ", data starts at row " & evalStartRow & vbCrLf
            Exit For
        End If
    Next i
    
    If evalStartRow = 0 Then
        MsgBox "ERROR: Cannot find 'Overall Status by Op Code' section!" & vbCrLf & vbCrLf & _
               "Evaluation Results sheet structure:" & vbCrLf & _
               "Expected: Row with text 'Overall Status by Op Code'" & vbCrLf & _
               "Next row: Column headers (Op Code, Operation, Overall Status...)" & vbCrLf & _
               "Following rows: Data", vbCritical, "Section Not Found"
        Exit Sub
    End If
    
    ' Step 3: Find "Operation Mode Summary" section
    debugMsg = debugMsg & vbCrLf & "Step 3: Locating 'Operation Mode Summary' section..." & vbCrLf
    summaryStartRow = 0
    
    For i = evalStartRow To evalLastRow
        If InStr(1, wsEval.Cells(i, 1).Value, "Operation Mode Summary", vbTextCompare) > 0 Then
            summaryStartRow = i + 2 ' Data starts 2 rows after header
            debugMsg = debugMsg & "✓ Found at row " & i & ", data starts at row " & summaryStartRow & vbCrLf
            Exit For
        End If
    Next i
    
    If summaryStartRow = 0 Then
        debugMsg = debugMsg & "⚠ 'Operation Mode Summary' section not found" & vbCrLf
    End If
    
    ' Step 4: Find "status" column in HeatMap Sheet
    debugMsg = debugMsg & vbCrLf & "Step 4: Finding 'status' column in HeatMap Sheet..." & vbCrLf
    statusCol = 0
    Dim headerRow As Long
    headerRow = 1 ' Assuming headers are in row 1
    
    ' Search for "status" column (case-insensitive)
    For j = 1 To 20 ' Check first 20 columns
        If InStr(1, wsHeatMap.Cells(headerRow, j).Value, "status", vbTextCompare) > 0 Then
            statusCol = j
            debugMsg = debugMsg & "✓ Found 'status' column at column " & j & " (" & Split(Cells(1, j).Address, "$")(1) & ")" & vbCrLf
            Exit For
        End If
    Next j
    
    If statusCol = 0 Then
        MsgBox "ERROR: Cannot find 'status' column in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Please ensure HeatMap Sheet has a column header containing 'status'" & vbCrLf & _
               "(checked first 20 columns in row 1)", vbCritical, "Column Not Found"
        Exit Sub
    End If
    
    heatMapLastRow = wsHeatMap.Cells(wsHeatMap.Rows.Count, 1).End(xlUp).Row
    debugMsg = debugMsg & "✓ HeatMap has " & heatMapLastRow & " rows" & vbCrLf & vbCrLf
    
    ' Step 5: Process sub-operations from "Overall Status by Op Code"
    debugMsg = debugMsg & "Step 5: Processing sub-operations..." & vbCrLf
    Dim subOpCount As Long
    subOpCount = 0
    
    For i = evalStartRow To evalLastRow
        opCode = Trim(wsEval.Cells(i, 1).Value)
        
        ' Stop if we hit Operation Mode Summary or empty row
        If opCode = "" Or InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
            Exit For
        End If
        
        ' Get status from column C (Overall Status)
        status = Trim(wsEval.Cells(i, 3).Value)
        
        If opCode <> "" And status <> "" Then
            ' Find matching operation in HeatMap
            For j = 2 To heatMapLastRow ' Start from row 2 (skip header)
                If Trim(wsHeatMap.Cells(j, 1).Value) = opCode Then
                    ' Update status column
                    wsHeatMap.Cells(j, statusCol).Value = GetStatusSymbol(status)
                    wsHeatMap.Cells(j, statusCol).Font.Name = "Wingdings"
                    wsHeatMap.Cells(j, statusCol).Font.Size = 14
                    wsHeatMap.Cells(j, statusCol).Font.Color = GetStatusColor(status)
                    
                    updateCount = updateCount + 1
                    subOpCount = subOpCount + 1
                    Exit For
                End If
            Next j
        End If
    Next i
    
    debugMsg = debugMsg & "✓ Updated " & subOpCount & " sub-operations" & vbCrLf
    
    ' Step 6: Process parent operations from "Operation Mode Summary"
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & vbCrLf & "Step 6: Processing parent operations..." & vbCrLf
        Dim parentOpCount As Long
        parentOpCount = 0
        
        For i = summaryStartRow To evalLastRow
            opCode = Trim(wsEval.Cells(i, 6).Value) ' Op Code is in column F
            
            If opCode = "" Then Exit For
            
            ' Get status from column I (Final Status)
            status = Trim(wsEval.Cells(i, 9).Value)
            
            If opCode <> "" And status <> "" Then
                ' Find matching operation in HeatMap
                For j = 2 To heatMapLastRow
                    If Trim(wsHeatMap.Cells(j, 1).Value) = opCode Then
                        ' Update status column
                        wsHeatMap.Cells(j, statusCol).Value = GetStatusSymbol(status)
                        wsHeatMap.Cells(j, statusCol).Font.Name = "Wingdings"
                        wsHeatMap.Cells(j, statusCol).Font.Size = 14
                        wsHeatMap.Cells(j, statusCol).Font.Color = GetStatusColor(status)
                        
                        updateCount = updateCount + 1
                        parentOpCount = parentOpCount + 1
                        Exit For
                    End If
                Next j
            End If
        Next i
        
        debugMsg = debugMsg & "✓ Updated " & parentOpCount & " parent operations" & vbCrLf
    End If
    
    ' Show results
    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 2)
    
    debugMsg = debugMsg & vbCrLf & "════════════════════════════════" & vbCrLf
    debugMsg = debugMsg & "COMPLETE!" & vbCrLf
    debugMsg = debugMsg & "Total operations updated: " & updateCount & vbCrLf
    debugMsg = debugMsg & "Time taken: " & elapsed & " seconds" & vbCrLf
    debugMsg = debugMsg & "════════════════════════════════"
    
    MsgBox debugMsg, vbInformation, "HeatMap Status Update Complete"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in UpdateHeatMapStatus:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Debug info:" & vbCrLf & debugMsg, vbCritical, "Error"
End Sub

' Helper function to get status symbol
Private Function GetStatusSymbol(status As String) As String
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusSymbol = "l" ' Filled circle in Wingdings
        Case "YELLOW"
            GetStatusSymbol = "l"
        Case "GREEN"
            GetStatusSymbol = "l"
        Case Else
            GetStatusSymbol = "l" ' Gray for N/A
    End Select
End Function

' Helper function to get status color
Private Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0) ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 255, 0) ' Yellow
        Case "GREEN"
            GetStatusColor = RGB(0, 255, 0) ' Green
        Case Else
            GetStatusColor = RGB(128, 128, 128) ' Gray for N/A
    End Select
End Function

' Create button on HeatMap Sheet
Public Sub CreateUpdateButton()
    On Error Resume Next
    
    Dim ws As Worksheet
    Dim btn As Button
    
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    
    If ws Is Nothing Then
        MsgBox "'HeatMap Sheet' not found!", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if present
    ws.Buttons.Delete
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    btn.OnAction = "UpdateHeatMapStatus"
    btn.Caption = "Update HeatMap Status"
    btn.Font.Bold = True
    btn.Font.Size = 12
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the 'Update HeatMap Status' button to transfer evaluation results.", _
           vbInformation, "Button Created"
End Sub
