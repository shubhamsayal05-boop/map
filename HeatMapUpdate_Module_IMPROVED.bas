Attribute VB_Name = "HeatMapUpdate_Improved"
' ====================================================================
' Module: HeatMapUpdate_Improved
' Purpose: Transfer evaluation results to HeatMap Sheet with enhanced debugging
' Version: 2.0 - Improved debugging and error messages
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
    Dim debugInfo As String
    Dim evalOpCount As Long
    Dim heatMapOpCount As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalOpCount = 0
    heatMapOpCount = 0
    debugInfo = ""
    
    ' Step 1: Verify sheets exist
    debugInfo = "=== HEATMAP UPDATE DEBUG INFO ===" & vbCrLf & vbCrLf
    debugInfo = debugInfo & "Step 1: Checking sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: Cannot find 'Evaluation Results' sheet!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: Cannot find 'HeatMap Sheet'!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing sheets..."
    
    ' Step 2: Analyze Evaluation Results structure
    debugInfo = debugInfo & "Step 2: Analyzing Evaluation Results sheet..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Total rows: " & lastRowEval & vbCrLf
    
    ' Find sections
    Dim overallRow As Long, summaryRow As Long
    overallRow = FindSectionRow(wsEval, "Overall Status by Op Code")
    summaryRow = FindSectionRow(wsEval, "Operation Mode Summary")
    
    If overallRow > 0 Then
        debugInfo = debugInfo & "  ✓ Found 'Overall Status by Op Code' at row " & overallRow & vbCrLf
    Else
        debugInfo = debugInfo & "  ✗ 'Overall Status by Op Code' NOT FOUND" & vbCrLf
    End If
    
    If summaryRow > 0 Then
        debugInfo = debugInfo & "  ✓ Found 'Operation Mode Summary' at row " & summaryRow & vbCrLf
    Else
        debugInfo = debugInfo & "  ✗ 'Operation Mode Summary' NOT FOUND" & vbCrLf
    End If
    debugInfo = debugInfo & vbCrLf
    
    ' Step 3: Analyze HeatMap Sheet structure
    debugInfo = debugInfo & "Step 3: Analyzing HeatMap Sheet..." & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Total rows: " & lastRowHeatMap & vbCrLf
    
    ' Find status column in HeatMap
    Dim statusCol As Long
    statusCol = FindColumnInRow(wsHeatMap, 1, "status")
    If statusCol = 0 Then
        statusCol = FindColumnInRow(wsHeatMap, 2, "status")
    End If
    If statusCol = 0 Then
        statusCol = FindColumnInRow(wsHeatMap, 3, "status")
    End If
    
    If statusCol > 0 Then
        debugInfo = debugInfo & "  ✓ Found 'status' column at column " & statusCol & " (" & ColumnLetter(statusCol) & ")" & vbCrLf
    Else
        debugInfo = debugInfo & "  ✗ 'status' column NOT FOUND in first 3 rows" & vbCrLf
        debugInfo = debugInfo & "  Row 1 headers: " & GetRowHeaders(wsHeatMap, 1, 10) & vbCrLf
    End If
    
    ' Count operations in HeatMap
    For i = 2 To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
            heatMapOpCount = heatMapOpCount + 1
        End If
    Next i
    debugInfo = debugInfo & "  Operation codes found: " & heatMapOpCount & vbCrLf & vbCrLf
    
    ' Step 4: Process Overall Status by Op Code section
    If overallRow > 0 Then
        debugInfo = debugInfo & "Step 4: Processing 'Overall Status by Op Code' section..." & vbCrLf
        Application.StatusBar = "Processing Overall Status section..."
        
        ' Find column with Final Status
        Dim finalStatusCol As Long
        finalStatusCol = FindColumnInRow(wsEval, overallRow + 1, "Final Status")
        If finalStatusCol = 0 Then
            finalStatusCol = FindColumnInRow(wsEval, overallRow + 1, "Overall Status")
        End If
        
        If finalStatusCol > 0 Then
            debugInfo = debugInfo & "  Status column found at: " & finalStatusCol & " (" & ColumnLetter(finalStatusCol) & ")" & vbCrLf
            
            ' Process this section
            For i = overallRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                ' Stop if we hit next section or empty
                If opCode = "" Or InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                    Exit For
                End If
                
                If IsNumeric(opCode) And Len(opCode) = 8 Then
                    evalOpCount = evalOpCount + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, finalStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        ' Update in HeatMap
                        If statusCol > 0 Then
                            If UpdateOperationInHeatMap(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap) Then
                                updatedCount = updatedCount + 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            debugInfo = debugInfo & "  Operations processed: " & evalOpCount & vbCrLf
            debugInfo = debugInfo & "  Operations updated: " & updatedCount & vbCrLf & vbCrLf
        Else
            debugInfo = debugInfo & "  ✗ Could not find Final Status column" & vbCrLf & vbCrLf
        End If
    End If
    
    ' Step 5: Process Operation Mode Summary section
    If summaryRow > 0 Then
        debugInfo = debugInfo & "Step 5: Processing 'Operation Mode Summary' section..." & vbCrLf
        Application.StatusBar = "Processing Operation Mode Summary..."
        
        ' Find columns in summary section
        Dim opCodeColSummary As Long, statusColSummary As Long
        opCodeColSummary = FindColumnInRow(wsEval, summaryRow + 1, "Op Code")
        statusColSummary = FindColumnInRow(wsEval, summaryRow + 1, "Final Status")
        
        If opCodeColSummary > 0 And statusColSummary > 0 Then
            debugInfo = debugInfo & "  Op Code column: " & opCodeColSummary & ", Status column: " & statusColSummary & vbCrLf
            
            Dim summaryCount As Long
            summaryCount = 0
            
            ' Process summary section
            For i = summaryRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, opCodeColSummary).Value))
                
                If opCode = "" Then Exit For
                
                If IsNumeric(opCode) And Len(opCode) = 8 Then
                    summaryCount = summaryCount + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColSummary).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        If statusCol > 0 Then
                            If UpdateOperationInHeatMap(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap) Then
                                updatedCount = updatedCount + 1
                            End If
                        End If
                    End If
                End If
            Next i
            
            debugInfo = debugInfo & "  Parent operations processed: " & summaryCount & vbCrLf & vbCrLf
        Else
            debugInfo = debugInfo & "  ✗ Could not find columns in summary section" & vbCrLf & vbCrLf
        End If
    End If
    
    ' Cleanup
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show results
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    debugInfo = debugInfo & "=== SUMMARY ===" & vbCrLf
    debugInfo = debugInfo & "Operations found in Evaluation: " & evalOpCount & vbCrLf
    debugInfo = debugInfo & "Operations found in HeatMap: " & heatMapOpCount & vbCrLf
    debugInfo = debugInfo & "Operations updated: " & updatedCount & vbCrLf
    debugInfo = debugInfo & "Time taken: " & Format(elapsedTime, "0.00") & " seconds"
    
    If updatedCount = 0 Then
        MsgBox debugInfo & vbCrLf & vbCrLf & _
               "⚠️ WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "1. Evaluation has been run" & vbCrLf & _
               "2. Op Codes match between sheets" & vbCrLf & _
               "3. Status column exists in HeatMap", _
               vbExclamation, "HeatMap Update Complete - No Updates"
    Else
        MsgBox debugInfo, vbInformation, "HeatMap Update Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error updating HeatMap:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           debugInfo, _
           vbCritical, "Error"
End Sub

' Helper function to find a section by name
Private Function FindSectionRow(ws As Worksheet, sectionName As String) As Long
    Dim i As Long
    Dim lastRow As Long
    
    FindSectionRow = 0
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, 1).Value), sectionName, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find a column by header name in a specific row
Private Function FindColumnInRow(ws As Worksheet, rowNum As Long, headerName As String) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindColumnInRow = 0
    
    For col = 1 To 50  ' Search first 50 columns
        cellValue = Trim(CStr(ws.Cells(rowNum, col).Value))
        If InStr(1, cellValue, headerName, vbTextCompare) > 0 Then
            FindColumnInRow = col
            Exit Function
        End If
    Next col
End Function

' Helper function to update an operation in HeatMap
Private Function UpdateOperationInHeatMap(ws As Worksheet, opCode As String, status As String, statusCol As Long, lastRow As Long) As Boolean
    Dim i As Long
    Dim hmOpCode As String
    
    UpdateOperationInHeatMap = False
    
    For i = 2 To lastRow
        hmOpCode = Trim(CStr(ws.Cells(i, 1).Value))
        If hmOpCode = opCode Then
            ' Found matching operation - update status
            ws.Cells(i, statusCol).Value = GetStatusDot(status)
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Size = 14
            ws.Cells(i, statusCol).Font.Color = GetStatusColor(status)
            UpdateOperationInHeatMap = True
            Exit Function
        End If
    Next i
End Function

' Helper function to get status dot character
Private Function GetStatusDot(status As String) As String
    GetStatusDot = "l"  ' Wingdings filled circle
End Function

' Helper function to get status color
Private Function GetStatusColor(status As String) As Long
    Select Case UCase(Trim(status))
        Case "RED"
            GetStatusColor = RGB(255, 0, 0)      ' Red
        Case "YELLOW"
            GetStatusColor = RGB(255, 192, 0)    ' Yellow
        Case "GREEN"
            GetStatusColor = RGB(0, 176, 80)     ' Green
        Case Else
            GetStatusColor = RGB(128, 128, 128)  ' Gray for N/A
    End Select
End Function

' Helper function to get column letter from number
Private Function ColumnLetter(col As Long) As String
    Dim dividend As Long
    Dim modulo As Long
    Dim result As String
    
    dividend = col
    Do While dividend > 0
        modulo = (dividend - 1) Mod 26
        result = Chr(65 + modulo) & result
        dividend = (dividend - modulo) \ 26
    Loop
    
    ColumnLetter = result
End Function

' Helper function to list all sheets
Private Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim result As String
    
    result = ""
    For Each ws In ThisWorkbook.Worksheets
        result = result & "  - " & ws.Name & vbCrLf
    Next ws
    
    ListAllSheets = result
End Function

' Helper function to get row headers for debugging
Private Function GetRowHeaders(ws As Worksheet, rowNum As Long, numCols As Long) As String
    Dim col As Long
    Dim result As String
    
    result = ""
    For col = 1 To numCols
        If ws.Cells(rowNum, col).Value <> "" Then
            result = result & ColumnLetter(col) & ":" & ws.Cells(rowNum, col).Value & " | "
        End If
    Next col
    
    GetRowHeaders = result
End Function

' Create Update Button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        MsgBox "Cannot find 'HeatMap Sheet'!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Remove existing button if it exists
    btnName = "UpdateHeatMapBtn"
    ws.Buttons(btnName).Delete
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = btnName
    btn.Caption = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to transfer evaluation results.", _
           vbInformation, "Button Created"
    
    On Error GoTo 0
End Sub
