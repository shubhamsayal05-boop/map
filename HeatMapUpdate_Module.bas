Attribute VB_Name = "HeatMapUpdate"
' ====================================================================
' Module: HeatMapUpdate
' Purpose: Transfer evaluation results to HeatMap Sheet status column
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
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    
    ' Get worksheets
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    
    ' Show progress message
    Application.ScreenUpdating = False
    Application.StatusBar = "Updating HeatMap statuses..."
    
    ' Find last rows
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    ' Loop through evaluation results (sub-operations first)
    ' Start from row 2 (skip header)
    For i = 2 To lastRowEval
        opCode = Trim(CStr(wsEval.Cells(i, 1).Value)) ' Column A: Op Code
        
        If opCode <> "" And IsNumeric(opCode) Then
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
    
    ' Now update parent operation modes from Operation Mode Summary section
    ' Find the "Operation Mode Summary" section in Evaluation Results
    Dim summaryStartRow As Long
    summaryStartRow = FindSummarySection(wsEval, lastRowEval)
    
    If summaryStartRow > 0 Then
        ' Loop through Operation Mode Summary rows
        For i = summaryStartRow + 1 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 6).Value)) ' Column F in summary: Op Code
            
            If opCode = "" Or Not IsNumeric(opCode) Then Exit For
            
            ' Get Final Status from column I (9th column) in summary section
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, 9).Value)))
            
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" Then
                ' Find this operation in HeatMap and update
                If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap) Then
                    updatedCount = updatedCount + 1
                End If
            End If
        Next i
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show completion message
    MsgBox "HeatMap updated successfully!" & vbCrLf & vbCrLf & _
           "Operations updated: " & updatedCount & vbCrLf & _
           "Time taken: " & Format(Timer - startTime, "0.0") & " seconds", _
           vbInformation, "Update Complete"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error updating HeatMap: " & Err.Description, vbCritical, "Update Error"
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
