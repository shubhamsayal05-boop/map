Attribute VB_Name = "HeatMapUpdate"
' ====================================================================
' Module: HeatMapUpdate
' Purpose: Transfer evaluation results to HeatMap Sheet status column
' ====================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results with enhanced debugging
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
    Dim evalFound As Long, heatMapFound As Long
    Dim statusColEval As Long, statusColSummary As Long
    Dim opCodeColSummary As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalFound = 0
    heatMapFound = 0
    debugMsg = ""
    
    ' Get worksheets with error checking
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Please ensure the sheet is named exactly 'Evaluation Results'", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Please ensure the sheet is named exactly 'HeatMap Sheet'", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    On Error GoTo ErrorHandler
    
    ' Show progress message
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing sheets..."
    
    ' Find last rows
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = "Sheet Analysis:" & vbCrLf & _
               "- Evaluation Results: " & lastRowEval & " rows" & vbCrLf & _
               "- HeatMap Sheet: " & lastRowHeatMap & " rows" & vbCrLf & vbCrLf
    
    ' Find "Overall Status by Op Code" section and determine status column
    Dim overallStartRow As Long
    overallStartRow = FindSectionRow(wsEval, "Overall Status by Op Code", lastRowEval)
    
    If overallStartRow > 0 Then
        debugMsg = debugMsg & "'Overall Status by Op Code' found at row " & overallStartRow & vbCrLf
        ' Find Final Status column in header row (overallStartRow + 1)
        statusColEval = FindColumnByHeader(wsEval, overallStartRow + 1, "Final Status")
        debugMsg = debugMsg & "Final Status column: " & statusColEval & vbCrLf & vbCrLf
        
        ' Loop through this section
        Application.StatusBar = "Processing Overall Status section..."
        For i = overallStartRow + 2 To lastRowEval ' Skip section title and header
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value)) ' Column A: Op Code
            
            ' Stop if we hit the next section
            If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            
            If opCode <> "" And IsNumeric(opCode) Then
                evalFound = evalFound + 1
                ' Get Final Status
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColEval).Value)))
                
                ' Skip if no status or header
                If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                    ' Find this operation in HeatMap and update
                    If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap) Then
                        updatedCount = updatedCount + 1
                    End If
                End If
            End If
        Next i
    Else
        debugMsg = debugMsg & "WARNING: 'Overall Status by Op Code' section NOT found!" & vbCrLf & vbCrLf
    End If
    
    ' Now update parent operation modes from Operation Mode Summary section
    Application.StatusBar = "Processing Operation Mode Summary..."
    Dim summaryStartRow As Long
    summaryStartRow = FindSectionRow(wsEval, "Operation Mode Summary", lastRowEval)
    
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "'Operation Mode Summary' found at row " & summaryStartRow & vbCrLf
        ' Find columns in header row
        opCodeColSummary = FindColumnByHeader(wsEval, summaryStartRow + 1, "Op Code")
        statusColSummary = FindColumnByHeader(wsEval, summaryStartRow + 1, "Final Status")
        debugMsg = debugMsg & "Op Code column: " & opCodeColSummary & ", Final Status column: " & statusColSummary & vbCrLf & vbCrLf
        
        ' Loop through Operation Mode Summary rows
        For i = summaryStartRow + 2 To lastRowEval ' Skip section title and header
            opCode = Trim(CStr(wsEval.Cells(i, opCodeColSummary).Value))
            
            If opCode = "" Or Not IsNumeric(opCode) Then Exit For
            
            evalFound = evalFound + 1
            ' Get Final Status
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColSummary).Value)))
            
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                ' Find this operation in HeatMap and update
                If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap) Then
                    updatedCount = updatedCount + 1
                End If
            End If
        Next i
    Else
        debugMsg = debugMsg & "WARNING: 'Operation Mode Summary' section NOT found!" & vbCrLf & vbCrLf
    End If
    
    ' Count operations in HeatMap
    For i = 1 To lastRowHeatMap
        If IsNumeric(Trim(CStr(wsHeatMap.Cells(i, 1).Value))) Then
            heatMapFound = heatMapFound + 1
        End If
    Next i
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show detailed results
    debugMsg = debugMsg & "Results:" & vbCrLf & _
               "- Operations found in Evaluation: " & evalFound & vbCrLf & _
               "- Operations found in HeatMap: " & heatMapFound & vbCrLf & _
               "- Successfully updated: " & updatedCount & vbCrLf & _
               "- Time taken: " & Format(Timer - startTime, "0.0") & " seconds"
    
    If updatedCount = 0 Then
        MsgBox "WARNING: No operations were updated!" & vbCrLf & vbCrLf & debugMsg, _
               vbExclamation, "Update Complete - Check Details"
    Else
        MsgBox "HeatMap updated successfully!" & vbCrLf & vbCrLf & debugMsg, _
               vbInformation, "Update Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error updating HeatMap:" & vbCrLf & vbCrLf & _
           "Error #" & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           debugMsg, _
           vbCritical, "Update Error"
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
    For i = 5 To lastRow ' Start from row 5 to skip headers
        heatMapOpCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        
        If heatMapOpCode = opCode Then
            ' Found the operation - update Current Status P1 (Column C or E depending on section)
            ' Try column C first (for Drivability section)
            Set statusCell = wsHeatMap.Cells(i, 3) ' Column C
            
            ' Check if this is a section header row (has "SET AS" or similar)
            If InStr(1, CStr(wsHeatMap.Cells(i, 3).Value), "SET AS", vbTextCompare) = 0 Then
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
        End If
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
        MsgBox "HeatMap Sheet not found!", vbCritical, "Error"
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
