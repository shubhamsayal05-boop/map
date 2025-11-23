Attribute VB_Name = "HeatMapUpdate_DEBUG"
' ====================================================================
' Module: HeatMapUpdate_DEBUG
' Purpose: Transfer evaluation results to HeatMap Sheet with debugging
' Version: 2.0 - Enhanced with detailed error messages
' ====================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results
Sub UpdateHeatMapStatus_DEBUG()
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
    Dim subOpCount As Long
    Dim parentOpCount As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    subOpCount = 0
    parentOpCount = 0
    debugMsg = "=== HeatMap Update Debug Report ===" & vbCrLf & vbCrLf
    
    ' Step 1: Check if sheets exist
    debugMsg = debugMsg & "Step 1: Checking sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "Error: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Please ensure:" & vbCrLf & _
               "1. You have run the evaluation (Alt+F8 → EvaluateAVLStatus)" & vbCrLf & _
               "2. The sheet is named exactly 'Evaluation Results'", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "  ✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "Error: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Please ensure the sheet is named exactly 'HeatMap Sheet'", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "  ✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    ' Show progress message
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing evaluation data..."
    
    ' Step 2: Find data ranges
    debugMsg = debugMsg & "Step 2: Analyzing data structure..." & vbCrLf
    
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = debugMsg & "  Evaluation Results last row: " & lastRowEval & vbCrLf
    debugMsg = debugMsg & "  HeatMap Sheet last row: " & lastRowHeatMap & vbCrLf & vbCrLf
    
    ' Step 3: Find "Overall Status by Op Code" section
    debugMsg = debugMsg & "Step 3: Looking for 'Overall Status by Op Code' section..." & vbCrLf
    
    Dim overallStatusRow As Long
    overallStatusRow = 0
    
    For i = 1 To lastRowEval
        Dim cellValue As String
        cellValue = Trim(CStr(wsEval.Cells(i, 1).Value))
        If InStr(1, cellValue, "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallStatusRow = i
            Exit For
        End If
    Next i
    
    If overallStatusRow = 0 Then
        debugMsg = debugMsg & "  ✗ 'Overall Status by Op Code' section NOT FOUND" & vbCrLf
        debugMsg = debugMsg & "  Searching first 20 rows of column A:" & vbCrLf
        For i = 1 To Application.Min(20, lastRowEval)
            debugMsg = debugMsg & "    Row " & i & ": " & wsEval.Cells(i, 1).Value & vbCrLf
        Next i
    Else
        debugMsg = debugMsg & "  ✓ Found at row: " & overallStatusRow & vbCrLf
        debugMsg = debugMsg & "  Headers in row " & (overallStatusRow + 1) & ":" & vbCrLf
        debugMsg = debugMsg & "    Col A: " & wsEval.Cells(overallStatusRow + 1, 1).Value & vbCrLf
        debugMsg = debugMsg & "    Col B: " & wsEval.Cells(overallStatusRow + 1, 2).Value & vbCrLf
        debugMsg = debugMsg & "    Col C: " & wsEval.Cells(overallStatusRow + 1, 3).Value & vbCrLf
        
        ' Process sub-operations from "Overall Status by Op Code" section
        Application.StatusBar = "Processing sub-operations..."
        
        For i = overallStatusRow + 2 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value)) ' Column A: Op Code
            
            ' Stop if we hit empty or next section
            If opCode = "" Or InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            
            If IsNumeric(opCode) Then
                ' Get Overall Status from column C (3rd column)
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, 3).Value)))
                
                If finalStatus <> "" And finalStatus <> "OVERALL STATUS" Then
                    subOpCount = subOpCount + 1
                    ' Find this operation in HeatMap and update
                    If UpdateOperationStatus_DEBUG(wsHeatMap, opCode, finalStatus, lastRowHeatMap, debugMsg) Then
                        updatedCount = updatedCount + 1
                    End If
                End If
            End If
        Next i
        
        debugMsg = debugMsg & "  Sub-operations processed: " & subOpCount & vbCrLf & vbCrLf
    End If
    
    ' Step 4: Find "Operation Mode Summary" section
    debugMsg = debugMsg & "Step 4: Looking for 'Operation Mode Summary' section..." & vbCrLf
    
    Dim summaryStartRow As Long
    summaryStartRow = 0
    
    For i = 1 To lastRowEval
        cellValue = Trim(CStr(wsEval.Cells(i, 1).Value))
        If InStr(1, cellValue, "Operation Mode Summary", vbTextCompare) > 0 Then
            summaryStartRow = i
            Exit For
        End If
    Next i
    
    If summaryStartRow = 0 Then
        debugMsg = debugMsg & "  ✗ 'Operation Mode Summary' section NOT FOUND" & vbCrLf
    Else
        debugMsg = debugMsg & "  ✓ Found at row: " & summaryStartRow & vbCrLf
        debugMsg = debugMsg & "  Headers in row " & (summaryStartRow + 1) & ":" & vbCrLf
        
        ' Check which column has Op Code - it should be column F (6th)
        For j = 1 To 10
            Dim headerVal As String
            headerVal = Trim(CStr(wsEval.Cells(summaryStartRow + 1, j).Value))
            If InStr(1, headerVal, "Op Code", vbTextCompare) > 0 Then
                debugMsg = debugMsg & "    Col " & Chr(64 + j) & " (" & j & "): " & headerVal & " ← Op Code column" & vbCrLf
            ElseIf InStr(1, headerVal, "Final Status", vbTextCompare) > 0 Then
                debugMsg = debugMsg & "    Col " & Chr(64 + j) & " (" & j & "): " & headerVal & " ← Final Status column" & vbCrLf
            End If
        Next j
        
        ' Process parent operations from "Operation Mode Summary" section
        Application.StatusBar = "Processing parent operations..."
        
        For i = summaryStartRow + 2 To lastRowEval
            ' Op Code is in column F (6th column) in summary section
            opCode = Trim(CStr(wsEval.Cells(i, 6).Value))
            
            If opCode = "" Or Not IsNumeric(opCode) Then Exit For
            
            ' Final Status is in column I (9th column) in summary section
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, 9).Value)))
            
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" Then
                parentOpCount = parentOpCount + 1
                ' Find this operation in HeatMap and update
                If UpdateOperationStatus_DEBUG(wsHeatMap, opCode, finalStatus, lastRowHeatMap, debugMsg) Then
                    updatedCount = updatedCount + 1
                End If
            End If
        Next i
        
        debugMsg = debugMsg & "  Parent operations processed: " & parentOpCount & vbCrLf & vbCrLf
    End If
    
    ' Step 5: Check HeatMap structure
    debugMsg = debugMsg & "Step 5: HeatMap Sheet structure..." & vbCrLf
    debugMsg = debugMsg & "  Column A header: " & wsHeatMap.Cells(1, 1).Value & vbCrLf
    
    ' Find Status column in HeatMap
    Dim statusCol As Long
    statusCol = 0
    For j = 1 To 20
        cellValue = Trim(CStr(wsHeatMap.Cells(1, j).Value))
        If InStr(1, cellValue, "Status", vbTextCompare) > 0 Then
            statusCol = j
            debugMsg = debugMsg & "  Found Status column at: " & Chr(64 + j) & " (" & j & "): " & cellValue & vbCrLf
            Exit For
        End If
    Next j
    
    If statusCol = 0 Then
        debugMsg = debugMsg & "  ✗ Status column NOT FOUND in row 1" & vbCrLf
    End If
    
    debugMsg = debugMsg & vbCrLf & "=== SUMMARY ===" & vbCrLf
    debugMsg = debugMsg & "Sub-operations found: " & subOpCount & vbCrLf
    debugMsg = debugMsg & "Parent operations found: " & parentOpCount & vbCrLf
    debugMsg = debugMsg & "Total operations updated: " & updatedCount & vbCrLf
    debugMsg = debugMsg & "Time taken: " & Format(Timer - startTime, "0.0") & " seconds" & vbCrLf
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show debug report
    MsgBox debugMsg, vbInformation, "HeatMap Update Debug Report"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error updating HeatMap: " & Err.Description & vbCrLf & vbCrLf & debugMsg, vbCritical, "Update Error"
End Sub

' Update a single operation status in HeatMap Sheet
Private Function UpdateOperationStatus_DEBUG(wsHeatMap As Worksheet, opCode As String, _
                                              finalStatus As String, lastRow As Long, _
                                              ByRef debugMsg As String) As Boolean
    Dim i As Long
    Dim heatMapOpCode As String
    Dim statusCol As Long
    
    UpdateOperationStatus_DEBUG = False
    statusCol = 3 ' Default to column C (Current Status P1)
    
    ' Search for this Op Code in HeatMap Sheet column A
    For i = 2 To lastRow
        heatMapOpCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        
        If heatMapOpCode = opCode Then
            ' Found matching operation - update status in column C
            Dim statusDot As String
            Dim dotColor As Long
            
            Select Case finalStatus
                Case "RED"
                    statusDot = "●"
                    dotColor = RGB(255, 0, 0) ' Red
                Case "YELLOW"
                    statusDot = "●"
                    dotColor = RGB(255, 255, 0) ' Yellow
                Case "GREEN"
                    statusDot = "●"
                    dotColor = RGB(0, 176, 80) ' Green
                Case Else
                    statusDot = "●"
                    dotColor = RGB(128, 128, 128) ' Gray for N/A
            End Select
            
            ' Update the cell
            wsHeatMap.Cells(i, statusCol).Value = statusDot
            wsHeatMap.Cells(i, statusCol).Font.Name = "Wingdings"
            wsHeatMap.Cells(i, statusCol).Font.Size = 14
            wsHeatMap.Cells(i, statusCol).Font.Color = dotColor
            wsHeatMap.Cells(i, statusCol).HorizontalAlignment = xlCenter
            
            UpdateOperationStatus_DEBUG = True
            Exit For
        End If
    Next i
    
    If Not UpdateOperationStatus_DEBUG Then
        debugMsg = debugMsg & "  ✗ Op Code " & opCode & " not found in HeatMap" & vbCrLf
    End If
End Function

' Create button on HeatMap Sheet
Sub CreateUpdateButton_DEBUG()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    
    If ws Is Nothing Then
        MsgBox "'HeatMap Sheet' not found!", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if it exists
    btnName = "btnUpdateHeatMap_DEBUG"
    ws.Buttons(btnName).Delete
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    btn.Name = btnName
    btn.Caption = "Update HeatMap Status (DEBUG)"
    btn.OnAction = "UpdateHeatMapStatus_DEBUG"
    btn.Font.Bold = True
    btn.Font.Size = 11
    
    MsgBox "Debug button created on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click it to see detailed diagnostic information.", vbInformation
End Sub
