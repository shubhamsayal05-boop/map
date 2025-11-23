Attribute VB_Name = "HeatMapUpdate_Enhanced"
' ====================================================================
' Module: HeatMapUpdate_Enhanced
' Purpose: Transfer evaluation results to HeatMap Sheet with debugging
' Version: 2.0 - Enhanced with comprehensive error messages
' ====================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results
Sub UpdateHeatMapStatus()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim lastRowEval As Long
    Dim lastRowHeatMap As Long
    Dim i As Long
    Dim opCode As String
    Dim finalStatus As String
    Dim updatedCount As Long
    Dim startTime As Double
    Dim statusCol As Long
    Dim overallStatusRow As Long
    Dim summaryRow As Long
    Dim debugMsg As String
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugMsg = "Debug Information:" & vbCrLf & vbCrLf
    
    ' Step 1: Check if sheets exist
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Please ensure the sheet name is exactly 'Evaluation Results'.", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Please ensure the sheet name is exactly 'HeatMap Sheet'.", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found 'HeatMap Sheet' sheet" & vbCrLf
    On Error GoTo ErrorHandler
    
    ' Step 2: Find the Status column in HeatMap sheet
    statusCol = FindStatusColumn(wsHeatMap)
    If statusCol = 0 Then
        MsgBox "ERROR: Could not find 'Status' column in HeatMap Sheet!" & vbCrLf & vbCrLf & _
               "Please ensure there is a column header named 'Status' (or similar)." & vbCrLf & _
               "Searched row 1 for: Status, Current Status, P1, Current Status P1", _
               vbCritical, "Column Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found Status column at: " & Chr(64 + statusCol) & vbCrLf
    
    ' Step 3: Find sections in Evaluation Results
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = debugMsg & "✓ Evaluation Results has " & lastRowEval & " rows" & vbCrLf
    debugMsg = debugMsg & "✓ HeatMap Sheet has " & lastRowHeatMap & " rows" & vbCrLf & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    overallStatusRow = FindSection(wsEval, "Overall Status by Op Code", lastRowEval)
    If overallStatusRow = 0 Then
        debugMsg = debugMsg & "⚠ 'Overall Status by Op Code' section not found" & vbCrLf
    Else
        debugMsg = debugMsg & "✓ Found 'Overall Status by Op Code' at row " & overallStatusRow & vbCrLf
    End If
    
    ' Find "Operation Mode Summary" section
    summaryRow = FindSection(wsEval, "Operation Mode Summary", lastRowEval)
    If summaryRow = 0 Then
        debugMsg = debugMsg & "⚠ 'Operation Mode Summary' section not found" & vbCrLf
    Else
        debugMsg = debugMsg & "✓ Found 'Operation Mode Summary' at row " & summaryRow & vbCrLf
    End If
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Updating HeatMap statuses..."
    
    ' Step 4: Process Overall Status by Op Code section
    If overallStatusRow > 0 Then
        debugMsg = debugMsg & vbCrLf & "Processing 'Overall Status by Op Code' section..." & vbCrLf
        
        ' Find which columns have the data
        Dim opCodeCol As Long, statusColEval As Long
        opCodeCol = FindColumnInRow(wsEval, overallStatusRow + 1, "Op Code")
        statusColEval = FindColumnInRow(wsEval, overallStatusRow + 1, "Overall Status")
        
        If opCodeCol > 0 And statusColEval > 0 Then
            debugMsg = debugMsg & "  Op Code column: " & Chr(64 + opCodeCol) & vbCrLf
            debugMsg = debugMsg & "  Status column: " & Chr(64 + statusColEval) & vbCrLf
            
            ' Process rows after header
            For i = overallStatusRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, opCodeCol).Value))
                
                ' Stop at empty row or next section
                If opCode = "" Or InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                    Exit For
                End If
                
                If IsNumeric(opCode) Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColEval).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "OVERALL STATUS" Then
                        If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap, statusCol) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
            
            debugMsg = debugMsg & "  Processed sub-operations: " & updatedCount & vbCrLf
        End If
    End If
    
    ' Step 5: Process Operation Mode Summary section
    If summaryRow > 0 Then
        debugMsg = debugMsg & vbCrLf & "Processing 'Operation Mode Summary' section..." & vbCrLf
        
        Dim summaryOpCol As Long, summaryStatusCol As Long
        Dim beforeCount As Long
        beforeCount = updatedCount
        
        summaryOpCol = FindColumnInRow(wsEval, summaryRow + 1, "Op Code")
        summaryStatusCol = FindColumnInRow(wsEval, summaryRow + 1, "Final Status")
        
        If summaryOpCol > 0 And summaryStatusCol > 0 Then
            debugMsg = debugMsg & "  Op Code column: " & Chr(64 + summaryOpCol) & vbCrLf
            debugMsg = debugMsg & "  Status column: " & Chr(64 + summaryStatusCol) & vbCrLf
            
            For i = summaryRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, summaryOpCol).Value))
                
                If opCode = "" Or Not IsNumeric(opCode) Then Exit For
                
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, summaryStatusCol).Value)))
                
                If finalStatus <> "" And finalStatus <> "FINAL STATUS" Then
                    If UpdateOperationStatus(wsHeatMap, opCode, finalStatus, lastRowHeatMap, statusCol) Then
                        updatedCount = updatedCount + 1
                    End If
                End If
            Next i
            
            debugMsg = debugMsg & "  Processed parent operations: " & (updatedCount - beforeCount) & vbCrLf
        End If
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show results
    If updatedCount = 0 Then
        MsgBox debugMsg & vbCrLf & _
               "⚠ WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               "Possible reasons:" & vbCrLf & _
               "1. Evaluation has not been run yet" & vbCrLf & _
               "2. Operation codes don't match between sheets" & vbCrLf & _
               "3. Status columns contain no data", _
               vbExclamation, "No Updates"
    Else
        MsgBox "✓ HeatMap updated successfully!" & vbCrLf & vbCrLf & _
               "Operations updated: " & updatedCount & vbCrLf & _
               "Time taken: " & Format(Timer - startTime, "0.0") & " seconds" & vbCrLf & vbCrLf & _
               debugMsg, _
               vbInformation, "Update Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "Error at line: " & Erl & vbCrLf & vbCrLf & _
           debugMsg, _
           vbCritical, "Update Error"
End Sub

' Find a section header in the worksheet
Private Function FindSection(ws As Worksheet, sectionName As String, lastRow As Long) As Long
    Dim i As Long
    Dim cellValue As String
    
    FindSection = 0
    
    For i = 1 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 1).Value))
        If InStr(1, cellValue, sectionName, vbTextCompare) > 0 Then
            FindSection = i
            Exit Function
        End If
    Next i
End Function

' Find a column by header name in a specific row
Private Function FindColumnInRow(ws As Worksheet, row As Long, headerName As String) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindColumnInRow = 0
    
    For col = 1 To 20 ' Search first 20 columns
        cellValue = Trim(CStr(ws.Cells(row, col).Value))
        If InStr(1, cellValue, headerName, vbTextCompare) > 0 Then
            FindColumnInRow = col
            Exit Function
        End If
    Next col
End Function

' Find the Status column in HeatMap sheet
Private Function FindStatusColumn(ws As Worksheet) As Long
    Dim col As Long
    Dim headerValue As String
    Dim searchTerms As Variant
    Dim term As Variant
    
    FindStatusColumn = 0
    searchTerms = Array("Status", "Current Status", "P1", "Current Status P1")
    
    ' Search row 1 for status column
    For col = 1 To 20
        headerValue = Trim(UCase(CStr(ws.Cells(1, col).Value)))
        
        For Each term In searchTerms
            If InStr(1, headerValue, UCase(term), vbTextCompare) > 0 Then
                FindStatusColumn = col
                Exit Function
            End If
        Next term
    Next col
End Function

' Update a single operation's status in HeatMap
Private Function UpdateOperationStatus(ws As Worksheet, opCode As String, _
                                      status As String, lastRow As Long, _
                                      statusCol As Long) As Boolean
    Dim i As Long
    Dim heatMapCode As String
    Dim statusDot As String
    Dim statusColor As Long
    
    UpdateOperationStatus = False
    
    ' Find the operation in HeatMap sheet (Column A)
    For i = 2 To lastRow
        heatMapCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If heatMapCode = opCode Then
            ' Get colored dot based on status
            Select Case status
                Case "RED"
                    statusDot = "●"
                    statusColor = RGB(255, 0, 0) ' Red
                Case "YELLOW"
                    statusDot = "●"
                    statusColor = RGB(255, 255, 0) ' Yellow
                Case "GREEN"
                    statusDot = "●"
                    statusColor = RGB(0, 255, 0) ' Green
                Case Else
                    statusDot = "●"
                    statusColor = RGB(128, 128, 128) ' Gray
            End Select
            
            ' Update the status cell
            With ws.Cells(i, statusCol)
                .Value = statusDot
                .Font.Name = "Wingdings"
                .Font.Size = 14
                .Font.Color = statusColor
                .HorizontalAlignment = xlCenter
            End With
            
            UpdateOperationStatus = True
            Exit Function
        End If
    Next i
End Function

' Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if present
    btnName = "btnUpdateHeatMap"
    ws.Buttons(btnName).Delete
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = btnName
    btn.OnAction = "UpdateHeatMapStatus"
    btn.Caption = "Update HeatMap Status"
    btn.Font.Bold = True
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button after running evaluation to transfer statuses.", _
           vbInformation, "Button Created"
End Sub

' Show debug information about sheet structure
Sub ShowSheetStructure()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim msg As String
    Dim i As Long
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    
    msg = "Sheet Structure Analysis:" & vbCrLf & vbCrLf
    
    If Not wsEval Is Nothing Then
        msg = msg & "✓ Evaluation Results sheet found" & vbCrLf
        msg = msg & "  First 10 rows in column A:" & vbCrLf
        For i = 1 To 10
            msg = msg & "    Row " & i & ": " & wsEval.Cells(i, 1).Value & vbCrLf
        Next i
    Else
        msg = msg & "✗ Evaluation Results sheet NOT found" & vbCrLf
    End If
    
    msg = msg & vbCrLf
    
    If Not wsHeatMap Is Nothing Then
        msg = msg & "✓ HeatMap Sheet found" & vbCrLf
        msg = msg & "  Headers in row 1:" & vbCrLf
        For i = 1 To 10
            If wsHeatMap.Cells(1, i).Value <> "" Then
                msg = msg & "    Col " & Chr(64 + i) & ": " & wsHeatMap.Cells(1, i).Value & vbCrLf
            End If
        Next i
    Else
        msg = msg & "✗ HeatMap Sheet NOT found" & vbCrLf
    End If
    
    MsgBox msg, vbInformation, "Sheet Structure"
End Sub
