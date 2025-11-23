Attribute VB_Name = "HeatMapUpdate"
' ====================================================================
' Module: HeatMapUpdate (Enhanced with Debug Messages)
' Purpose: Transfer evaluation results to HeatMap Sheet status column
' Version: 2.0 - With detailed diagnostics
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
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugMsg = ""
    
    ' Check if Evaluation Results sheet exists
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Please run the evaluation first (Alt+F8 -> EvaluateAVLStatus).", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Check if HeatMap Sheet exists
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Please ensure your workbook has a sheet named exactly 'HeatMap Sheet'.", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & "✓ Found both sheets" & vbCrLf
    
    ' Show progress message
    Application.ScreenUpdating = False
    Application.StatusBar = "Updating HeatMap statuses..."
    
    ' Find last rows
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = debugMsg & "✓ Evaluation Results: " & lastRowEval & " rows" & vbCrLf
    debugMsg = debugMsg & "✓ HeatMap Sheet: " & lastRowHeatMap & " rows" & vbCrLf & vbCrLf
    
    ' Check if evaluation has data
    If lastRowEval < 2 Then
        Application.ScreenUpdating = True
        Application.StatusBar = False
        MsgBox "ERROR: No data found in Evaluation Results sheet!" & vbCrLf & vbCrLf & _
               "The sheet appears to be empty. Please run the evaluation first.", _
               vbCritical, "No Data"
        Exit Sub
    End If
    
    ' Loop through evaluation results (sub-operations first)
    ' Start from row 2 (skip header)
    Dim subOpCount As Long
    subOpCount = 0
    
    For i = 2 To lastRowEval
        opCode = Trim(CStr(wsEval.Cells(i, 1).Value)) ' Column A: Op Code
        
        If opCode <> "" And IsNumeric(opCode) Then
            subOpCount = subOpCount + 1
            ' Get Final Status from column M (13th column)
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, 13).Value)))
            
            ' Skip if no status
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" Then
                ' Find this operation in HeatMap and update
                If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap) Then
                    updatedCount = updatedCount + 1
                End If
            End If
        End If
    Next i
    
    debugMsg = debugMsg & "Sub-operations found: " & subOpCount & vbCrLf
    debugMsg = debugMsg & "Sub-operations updated: " & updatedCount & vbCrLf & vbCrLf
    
    ' Now update parent operation modes from Operation Mode Summary section
    ' Find the "Operation Mode Summary" section in Evaluation Results
    Dim summaryStartRow As Long
    Dim parentOpCount As Long
    summaryStartRow = FindSummarySection(wsEval, lastRowEval)
    parentOpCount = 0
    
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "✓ Found Operation Mode Summary at row " & summaryStartRow & vbCrLf
        
        ' Loop through Operation Mode Summary rows
        For i = summaryStartRow + 1 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 6).Value)) ' Column F in summary: Op Code
            
            If opCode = "" Or Not IsNumeric(opCode) Then Exit For
            
            parentOpCount = parentOpCount + 1
            ' Get Final Status from column I (9th column) in summary section
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, 9).Value)))
            
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" Then
                ' Find this operation in HeatMap and update
                If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap) Then
                    updatedCount = updatedCount + 1
                End If
            End If
        Next i
        
        debugMsg = debugMsg & "Parent operations found: " & parentOpCount & vbCrLf
    Else
        debugMsg = debugMsg & "⚠ Warning: Operation Mode Summary section not found" & vbCrLf
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show completion message with diagnostics
    If updatedCount > 0 Then
        MsgBox "HeatMap updated successfully!" & vbCrLf & vbCrLf & _
               debugMsg & vbCrLf & _
               "TOTAL operations updated: " & updatedCount & vbCrLf & _
               "Time taken: " & Format(Timer - startTime, "0.0") & " seconds", _
               vbInformation, "Update Complete"
    Else
        MsgBox "WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               debugMsg & vbCrLf & _
               "Possible reasons:" & vbCrLf & _
               "1. Op Codes in HeatMap don't match those in Evaluation Results" & vbCrLf & _
               "2. Final Status column is empty in Evaluation Results" & vbCrLf & _
               "3. Sheet structure is different than expected" & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "- Evaluation Results Column A has Op Codes (e.g., 10101300)" & vbCrLf & _
               "- Evaluation Results Column M has Final Status" & vbCrLf & _
               "- HeatMap Sheet Column A has matching Op Codes", _
               vbExclamation, "No Updates Made"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error updating HeatMap: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl, vbCritical, "Update Error"
End Sub

' Find the Operation Mode Summary section in Evaluation Results
Private Function FindSummarySection(ws As Worksheet, lastRow As Long) As Long
    Dim i As Long
    FindSummarySection = 0
    
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
            FindSummarySection = i
            Exit Function
        End If
    Next i
End Function

' Update a specific operation's status in HeatMap
Private Function UpdateOperationStatus(wsHeatMap As Worksheet, opCode As String, _
                                      status As String, lastRow As Long) As Boolean
    Dim i As Long
    Dim heatMapOpCode As String
    Dim statusCell As Range
    Dim dotChar As String
    
    UpdateOperationStatus = False
    dotChar = ChrW(&H25CF) ' Unicode filled circle: ●
    
    ' Search for operation code in HeatMap (Column A)
    For i = 1 To lastRow ' Start from row 1 to be thorough
        heatMapOpCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        
        If heatMapOpCode = opCode Then
            ' Found the operation - update Current Status P1 (Column C)
            Set statusCell = wsHeatMap.Cells(i, 3) ' Column C: Current Status P1
            
            ' Skip if this is a header row
            Dim cellValue As String
            cellValue = UCase(Trim(CStr(statusCell.Value)))
            If InStr(1, cellValue, "SET AS", vbTextCompare) > 0 Or _
               InStr(1, cellValue, "USE CASE", vbTextCompare) > 0 Or _
               InStr(1, cellValue, "CURRENT STATUS", vbTextCompare) > 0 Then
                ' This is a header row, skip it
                GoTo ContinueLoop
            End If
            
            ' Clear existing content
            statusCell.ClearContents
            statusCell.Font.Size = 14
            statusCell.Font.Name = "Wingdings"
            
            ' Set the dot character
            statusCell.Value = "l" ' Wingdings character for filled circle
            
            ' Set color based on status
            Select Case status
                Case "RED"
                    statusCell.Font.Color = RGB(255, 0, 0) ' Red
                Case "YELLOW"
                    statusCell.Font.Color = RGB(255, 192, 0) ' Yellow/Orange
                Case "GREEN"
                    statusCell.Font.Color = RGB(0, 176, 80) ' Green
                Case Else
                    statusCell.Font.Color = RGB(166, 166, 166) ' Gray for N/A
            End Select
            
            ' Center the dot
            statusCell.HorizontalAlignment = xlCenter
            statusCell.VerticalAlignment = xlCenter
            
            UpdateOperationStatus = True
            Exit Function
        End If
        
ContinueLoop:
    Next i
End Function

' Create the "Update HeatMap Status" button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim wsHeatMap As Worksheet
    Dim btn As Button
    Dim btnExists As Boolean
    Dim obj As Object
    
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If wsHeatMap Is Nothing Then
        MsgBox "HeatMap Sheet not found!" & vbCrLf & vbCrLf & _
               "Please ensure your workbook has a sheet named exactly 'HeatMap Sheet'.", _
               vbCritical, "Error"
        Exit Sub
    End If
    
    ' Check if button already exists
    btnExists = False
    For Each obj In wsHeatMap.Buttons
        If obj.Name = "btnUpdateHeatMap" Or obj.Caption = "Update HeatMap Status" Then
            btnExists = True
            MsgBox "Button already exists on HeatMap Sheet!", vbInformation, "Button Exists"
            Exit Sub
        End If
    Next obj
    
    ' Create button
    Set btn = wsHeatMap.Buttons.Add(10, 10, 150, 30) ' Left, Top, Width, Height
    With btn
        .Name = "btnUpdateHeatMap"
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to transfer evaluation results to HeatMap.", _
           vbInformation, "Button Created"
End Sub

' Diagnostic function to check data structure
Sub DiagnoseHeatMapIssue()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim msg As String
    Dim i As Long
    
    msg = "=== DIAGNOSTIC REPORT ===" & vbCrLf & vbCrLf
    
    ' Check Evaluation Results
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        msg = msg & "❌ 'Evaluation Results' sheet NOT FOUND" & vbCrLf
    Else
        msg = msg & "✓ 'Evaluation Results' sheet found" & vbCrLf
        msg = msg & "  - Last row: " & wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row & vbCrLf
        msg = msg & "  - First Op Code (A2): " & wsEval.Cells(2, 1).Value & vbCrLf
        msg = msg & "  - First Final Status (M2): " & wsEval.Cells(2, 13).Value & vbCrLf
    End If
    
    ' Check HeatMap Sheet
    Set wsHeatMap = Nothing
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        msg = msg & "❌ 'HeatMap Sheet' NOT FOUND" & vbCrLf
    Else
        msg = msg & "✓ 'HeatMap Sheet' found" & vbCrLf
        msg = msg & "  - Last row: " & wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row & vbCrLf
        msg = msg & "  - Sample Op Codes:" & vbCrLf
        For i = 1 To 10
            If wsHeatMap.Cells(i, 1).Value <> "" Then
                msg = msg & "    Row " & i & ": " & wsHeatMap.Cells(i, 1).Value & vbCrLf
            End If
        Next i
    End If
    On Error GoTo 0
    
    MsgBox msg, vbInformation, "Diagnostic Report"
End Sub
