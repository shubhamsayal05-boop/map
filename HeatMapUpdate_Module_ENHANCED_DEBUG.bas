Attribute VB_Name = "HeatMapUpdate_Enhanced_Debug"
' ====================================================================
' Module: HeatMapUpdate_Enhanced_Debug
' Purpose: Transfer evaluation results to HeatMap Sheet with detailed debugging
' ====================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results with debug output
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
    Dim statusCol As Long
    Dim summaryStartRow As Long
    Dim debugMsg As String
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugMsg = ""
    
    ' Step 1: Get worksheets
    debugMsg = debugMsg & "Step 1: Getting worksheets..." & vbCrLf
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    debugMsg = debugMsg & "  ✓ Found Evaluation Results sheet" & vbCrLf
    debugMsg = debugMsg & "  ✓ Found HeatMap Sheet" & vbCrLf & vbCrLf
    
    ' Step 2: Find Status column in HeatMap Sheet
    debugMsg = debugMsg & "Step 2: Finding Status column in HeatMap Sheet..." & vbCrLf
    statusCol = FindStatusColumn(wsHeatMap)
    If statusCol = 0 Then
        MsgBox "Could not find 'Status' column in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Please ensure there is a column with 'Status' in the header.", vbCritical
        Exit Sub
    End If
    debugMsg = debugMsg & "  ✓ Status column found at column " & statusCol & " (" & _
               Split(Cells(1, statusCol).Address, "$")(1) & ")" & vbCrLf & vbCrLf
    
    ' Step 3: Find data ranges
    debugMsg = debugMsg & "Step 3: Finding data ranges..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "  Evaluation Results last row: " & lastRowEval & vbCrLf
    debugMsg = debugMsg & "  HeatMap Sheet last row: " & lastRowHeatMap & vbCrLf & vbCrLf
    
    ' Step 4: Find "Operation Mode Summary" section
    debugMsg = debugMsg & "Step 4: Finding sections in Evaluation Results..." & vbCrLf
    summaryStartRow = FindSummarySection(wsEval, lastRowEval)
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "  ✓ Found 'Operation Mode Summary' at row " & summaryStartRow & vbCrLf
    Else
        debugMsg = debugMsg & "  ⚠ 'Operation Mode Summary' section not found" & vbCrLf
    End If
    debugMsg = debugMsg & vbCrLf
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Updating HeatMap statuses..."
    
    ' Step 5: Process sub-operations (Overall Status by Op Code section)
    debugMsg = debugMsg & "Step 5: Processing sub-operations..." & vbCrLf
    Dim subOpsCount As Long: subOpsCount = 0
    
    For i = 2 To IIf(summaryStartRow > 0, summaryStartRow - 1, lastRowEval)
        opCode = Trim(CStr(wsEval.Cells(i, 1).Value)) ' Column A: Op Code
        
        If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
            ' Get Final Status from column M (13th column)
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, 13).Value)))
            
            ' Skip header rows
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                ' Find and update in HeatMap
                If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap, statusCol) Then
                    updatedCount = updatedCount + 1
                    subOpsCount = subOpsCount + 1
                End If
            End If
        End If
    Next i
    debugMsg = debugMsg & "  Processed " & subOpsCount & " sub-operations" & vbCrLf & vbCrLf
    
    ' Step 6: Process parent operations (Operation Mode Summary section)
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "Step 6: Processing parent operations..." & vbCrLf
        Dim parentOpsCount As Long: parentOpsCount = 0
        
        For i = summaryStartRow + 1 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 6).Value)) ' Column F: Op Code in summary
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                ' Get Final Status from column I (9th column in summary section)
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, 9).Value)))
                
                ' Skip header rows
                If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                    ' Find and update in HeatMap
                    If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap, statusCol) Then
                        updatedCount = updatedCount + 1
                        parentOpsCount = parentOpsCount + 1
                    End If
                End If
            End If
        Next i
        debugMsg = debugMsg & "  Processed " & parentOpsCount & " parent operations" & vbCrLf & vbCrLf
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show results
    Dim elapsedTime As Double
    elapsedTime = Round(Timer - startTime, 2)
    
    debugMsg = debugMsg & "========================================" & vbCrLf
    debugMsg = debugMsg & "SUMMARY" & vbCrLf
    debugMsg = debugMsg & "========================================" & vbCrLf
    debugMsg = debugMsg & "Total operations updated: " & updatedCount & vbCrLf
    debugMsg = debugMsg & "Time taken: " & elapsedTime & " seconds" & vbCrLf
    
    If updatedCount > 0 Then
        MsgBox debugMsg, vbInformation, "HeatMap Status Update Complete"
    Else
        MsgBox debugMsg & vbCrLf & vbCrLf & _
               "⚠ WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               "Possible issues:" & vbCrLf & _
               "1. Operation codes in Evaluation Results don't match HeatMap Sheet" & vbCrLf & _
               "2. All statuses are N/A or empty" & vbCrLf & _
               "3. Column structure is different than expected", vbExclamation, "No Updates"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error updating HeatMap status:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Debug info:" & vbCrLf & debugMsg, vbCritical
End Sub

' Find the Status column in HeatMap Sheet
Private Function FindStatusColumn(ws As Worksheet) As Long
    Dim col As Long
    Dim headerVal As String
    
    FindStatusColumn = 0
    
    ' Search first row for "Status" column
    For col = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        headerVal = Trim(UCase(CStr(ws.Cells(1, col).Value)))
        If InStr(headerVal, "STATUS") > 0 Then
            FindStatusColumn = col
            Exit Function
        End If
    Next col
End Function

' Find the Operation Mode Summary section
Private Function FindSummarySection(ws As Worksheet, lastRow As Long) As Long
    Dim i As Long
    Dim cellVal As String
    
    FindSummarySection = 0
    
    ' Search for "Operation Mode Summary" text
    For i = 1 To lastRow
        cellVal = Trim(CStr(ws.Cells(i, 1).Value))
        If InStr(1, cellVal, "Operation Mode Summary", vbTextCompare) > 0 Then
            FindSummarySection = i
            Exit Function
        End If
    Next i
End Function

' Update a single operation status in HeatMap
Private Function UpdateOperationStatus(ws As Worksheet, opCode As String, _
                                      status As String, lastRow As Long, statusCol As Long) As Boolean
    Dim i As Long
    Dim heatMapOpCode As String
    
    UpdateOperationStatus = False
    
    ' Search for matching operation code in HeatMap
    For i = 2 To lastRow ' Start from row 2 (skip header)
        heatMapOpCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If heatMapOpCode = opCode Then
            ' Set status with colored dot
            ws.Cells(i, statusCol).Value = GetStatusDot(status)
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Size = 14
            ws.Cells(i, statusCol).Font.Color = GetStatusColor(status)
            
            UpdateOperationStatus = True
            Exit Function
        End If
    Next i
End Function

' Get colored dot character for status
Private Function GetStatusDot(status As String) As String
    GetStatusDot = Chr(108) ' Filled circle in Wingdings font (●)
End Function

' Get color code for status
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
    Dim obj As Object
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Check if button already exists
    btnExists = False
    For Each obj In ws.Buttons
        If obj.Text = "Update HeatMap Status" Then
            btnExists = True
            MsgBox "Button already exists on HeatMap Sheet!", vbInformation
            Exit Sub
        End If
    Next obj
    
    ' Create button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Text = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update statuses after running evaluation.", vbInformation
End Sub
