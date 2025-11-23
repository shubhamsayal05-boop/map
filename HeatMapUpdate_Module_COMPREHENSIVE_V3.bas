Attribute VB_Name = "HeatMapUpdate_V3"
' ====================================================================
' Module: HeatMapUpdate_V3
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Version: 3.0 - Enhanced diagnostics and flexible sheet structure
' ====================================================================

Option Explicit

' Main function with comprehensive debugging
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
    Dim debugInfo As String
    Dim statusCol As Long
    Dim foundMatch As Boolean
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugInfo = "=== HEATMAP UPDATE DIAGNOSTIC ===" & vbCrLf & vbCrLf
    
    ' === STEP 1: Verify sheets exist ===
    debugInfo = debugInfo & "STEP 1: Checking sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & ListAllSheets(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "  ✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    
    ' === STEP 2: Analyze Evaluation Results sheet ===
    debugInfo = debugInfo & "STEP 2: Analyzing Evaluation Results..." & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Total rows: " & lastRowEval & vbCrLf
    
    ' Scan for key sections
    Dim overallRow As Long, summaryRow As Long
    overallRow = 0
    summaryRow = 0
    
    For i = 1 To lastRowEval
        Dim cellVal As String
        cellVal = Trim(CStr(wsEval.Cells(i, 1).Value))
        
        If InStr(1, cellVal, "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallRow = i
            debugInfo = debugInfo & "  ✓ Found 'Overall Status by Op Code' at row " & i & vbCrLf
        End If
        
        If InStr(1, cellVal, "Operation Mode Summary", vbTextCompare) > 0 Then
            summaryRow = i
            debugInfo = debugInfo & "  ✓ Found 'Operation Mode Summary' at row " & i & vbCrLf
        End If
        
        ' Stop after finding both sections
        If overallRow > 0 And summaryRow > 0 And i > summaryRow + 5 Then
            Exit For
        End If
    Next i
    debugInfo = debugInfo & vbCrLf
    
    If overallRow = 0 And summaryRow = 0 Then
        MsgBox "ERROR: Could not find required sections in Evaluation Results!" & vbCrLf & vbCrLf & _
               "Looking for:" & vbCrLf & _
               "- 'Overall Status by Op Code'" & vbCrLf & _
               "- 'Operation Mode Summary'" & vbCrLf & vbCrLf & _
               "Please check sheet structure.", vbCritical, "Section Not Found"
        Application.ScreenUpdating = True
        Exit Sub
    End If
    
    ' === STEP 3: Analyze HeatMap Sheet ===
    debugInfo = debugInfo & "STEP 3: Analyzing HeatMap Sheet..." & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "  Total rows: " & lastRowHeatMap & vbCrLf
    
    ' Find Status column (should be column with "Status" header)
    statusCol = 0
    For j = 1 To 10 ' Check first 10 columns
        Dim headerVal As String
        headerVal = Trim(UCase(CStr(wsHeatMap.Cells(1, j).Value)))
        If headerVal = "STATUS" Or headerVal = "CURRENT STATUS" Or headerVal = "CURRENT STATUS P1" Then
            statusCol = j
            debugInfo = debugInfo & "  ✓ Found Status column: " & j & " ('" & wsHeatMap.Cells(1, j).Value & "')" & vbCrLf
            Exit For
        End If
    Next j
    
    If statusCol = 0 Then
        ' Default to column B if not found
        statusCol = 2
        debugInfo = debugInfo & "  ! Status column not found, defaulting to column B" & vbCrLf
    End If
    debugInfo = debugInfo & vbCrLf
    
    ' === STEP 4: Process Overall Status section ===
    If overallRow > 0 Then
        debugInfo = debugInfo & "STEP 4: Processing 'Overall Status by Op Code'..." & vbCrLf
        Application.StatusBar = "Processing Overall Status section..."
        
        ' Find Final Status column
        Dim finalStatusCol As Long
        finalStatusCol = 0
        Dim headerRow As Long
        headerRow = overallRow + 1 ' Header is usually next row after section title
        
        For j = 1 To 20 ' Check first 20 columns
            Dim hdr As String
            hdr = Trim(UCase(CStr(wsEval.Cells(headerRow, j).Value)))
            If InStr(1, hdr, "FINAL STATUS", vbTextCompare) > 0 Or hdr = "OVERALL STATUS" Then
                finalStatusCol = j
                debugInfo = debugInfo & "  ✓ Found Final Status column: " & j & vbCrLf
                Exit For
            End If
        Next j
        
        If finalStatusCol = 0 Then
            debugInfo = debugInfo & "  ! Final Status column not found, trying common positions..." & vbCrLf
            ' Try common column positions
            If UCase(Trim(CStr(wsEval.Cells(headerRow, 3).Value))) Like "*STATUS*" Then
                finalStatusCol = 3
            ElseIf UCase(Trim(CStr(wsEval.Cells(headerRow, 4).Value))) Like "*STATUS*" Then
                finalStatusCol = 4
            Else
                finalStatusCol = 3 ' Default to column C
            End If
            debugInfo = debugInfo & "  Using column " & finalStatusCol & vbCrLf
        End If
        
        ' Process data rows
        Dim dataStartRow As Long
        dataStartRow = headerRow + 1
        Dim processedCount As Long
        processedCount = 0
        
        For i = dataStartRow To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Stop if we hit next section
            If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Or _
               InStr(1, opCode, "Accelerations", vbTextCompare) > 0 Or _
               opCode = "" Then
                If summaryRow > 0 And i >= summaryRow Then
                    Exit For
                End If
            End If
            
            ' Process if valid operation code
            If Len(opCode) >= 8 And IsNumeric(opCode) Then
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, finalStatusCol).Value)))
                processedCount = processedCount + 1
                
                If finalStatus <> "" And finalStatus <> "N/A" Then
                    ' Update in HeatMap
                    foundMatch = UpdateHeatMapOperation(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap)
                    If foundMatch Then
                        updatedCount = updatedCount + 1
                    End If
                End If
            End If
        Next i
        
        debugInfo = debugInfo & "  Processed " & processedCount & " operations" & vbCrLf
        debugInfo = debugInfo & "  Updated " & updatedCount & " in HeatMap" & vbCrLf & vbCrLf
    End If
    
    ' === STEP 5: Process Operation Mode Summary section ===
    If summaryRow > 0 Then
        debugInfo = debugInfo & "STEP 5: Processing 'Operation Mode Summary'..." & vbCrLf
        Application.StatusBar = "Processing Operation Mode Summary..."
        
        ' Find columns in summary section
        Dim summaryHeaderRow As Long
        summaryHeaderRow = summaryRow + 1
        Dim opCodeColSummary As Long, statusColSummary As Long
        opCodeColSummary = 0
        statusColSummary = 0
        
        For j = 1 To 20
            Dim sumHdr As String
            sumHdr = Trim(UCase(CStr(wsEval.Cells(summaryHeaderRow, j).Value)))
            If sumHdr = "OP CODE" Or sumHdr = "OPCODE" Then
                opCodeColSummary = j
            End If
            If InStr(1, sumHdr, "FINAL STATUS", vbTextCompare) > 0 Then
                statusColSummary = j
            End If
        Next j
        
        If opCodeColSummary = 0 Then opCodeColSummary = 1 ' Default to column A
        If statusColSummary = 0 Then statusColSummary = 3 ' Default to column C
        
        debugInfo = debugInfo & "  Op Code column: " & opCodeColSummary & vbCrLf
        debugInfo = debugInfo & "  Final Status column: " & statusColSummary & vbCrLf
        
        ' Process summary data
        Dim summaryProcessed As Long
        summaryProcessed = 0
        
        For i = summaryHeaderRow + 1 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, opCodeColSummary).Value))
            
            ' Stop if empty or end of data
            If opCode = "" Or Not IsNumeric(opCode) Then
                ' Check a few more rows in case of gaps
                If i > summaryHeaderRow + 50 Then Exit For
            Else
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColSummary).Value)))
                summaryProcessed = summaryProcessed + 1
                
                If finalStatus <> "" And finalStatus <> "N/A" Then
                    foundMatch = UpdateHeatMapOperation(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap)
                    If foundMatch Then
                        updatedCount = updatedCount + 1
                    End If
                End If
            End If
        Next i
        
        debugInfo = debugInfo & "  Processed " & summaryProcessed & " parent operations" & vbCrLf
        debugInfo = debugInfo & "  Total updated: " & updatedCount & vbCrLf & vbCrLf
    End If
    
    ' === STEP 6: Complete ===
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim elapsedTime As String
    elapsedTime = Format(Timer - startTime, "0.00")
    
    debugInfo = debugInfo & "=== COMPLETE ===" & vbCrLf
    debugInfo = debugInfo & "Operations updated: " & updatedCount & vbCrLf
    debugInfo = debugInfo & "Time taken: " & elapsedTime & " seconds"
    
    If updatedCount = 0 Then
        MsgBox debugInfo & vbCrLf & vbCrLf & _
               "WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               "Please check:" & vbCrLf & _
               "1. Evaluation has been run" & vbCrLf & _
               "2. Operation codes match between sheets" & vbCrLf & _
               "3. Status values are valid (RED/YELLOW/GREEN)", _
               vbExclamation, "Update Complete - No Changes"
    Else
        MsgBox "HeatMap Status Update Complete!" & vbCrLf & vbCrLf & _
               "Operations updated: " & updatedCount & vbCrLf & _
               "Time taken: " & elapsedTime & " seconds", _
               vbInformation, "Success"
    End If
    
    ' Show detailed debug info if requested
    Dim response As VbMsgBoxResult
    response = MsgBox("Would you like to see detailed diagnostic information?", vbQuestion + vbYesNo, "Diagnostics")
    If response = vbYes Then
        MsgBox debugInfo, vbInformation, "Diagnostic Information"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl, vbCritical, "Update Failed"
End Sub

' Helper function to update operation in HeatMap
Function UpdateHeatMapOperation(ws As Worksheet, opCode As String, status As String, statusCol As Long, lastRow As Long) As Boolean
    Dim i As Long
    Dim hmOpCode As String
    Dim dotChar As String
    Dim dotColor As Long
    
    UpdateHeatMapOperation = False
    
    ' Search for operation code in HeatMap
    For i = 2 To lastRow ' Start from row 2 (assuming row 1 is header)
        hmOpCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If hmOpCode = opCode Then
            ' Found matching operation - update status
            Select Case status
                Case "RED"
                    dotChar = "●"
                    dotColor = RGB(255, 0, 0) ' Red
                Case "YELLOW"
                    dotChar = "●"
                    dotColor = RGB(255, 255, 0) ' Yellow
                Case "GREEN"
                    dotChar = "●"
                    dotColor = RGB(0, 255, 0) ' Green
                Case Else
                    dotChar = "●"
                    dotColor = RGB(128, 128, 128) ' Gray for N/A
            End Select
            
            ' Update cell
            ws.Cells(i, statusCol).Value = dotChar
            ws.Cells(i, statusCol).Font.Name = "Wingdings"
            ws.Cells(i, statusCol).Font.Color = dotColor
            ws.Cells(i, statusCol).Font.Size = 14
            
            UpdateHeatMapOperation = True
            Exit Function
        End If
    Next i
End Function

' List all sheets in workbook
Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        sheetList = sheetList & "  - " & ws.Name & vbCrLf
    Next ws
    
    ListAllSheets = sheetList
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
    On Error GoTo 0
    
    ' Delete existing button if present
    btnName = "btnUpdateHeatMap"
    On Error Resume Next
    ws.Buttons(btnName).Delete
    On Error GoTo 0
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 180, 30)
    btn.Name = btnName
    btn.OnAction = "UpdateHeatMapStatus"
    btn.Caption = "Update HeatMap Status"
    btn.Font.Bold = True
    btn.Font.Size = 11
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click 'Update HeatMap Status' button after running evaluation.", _
           vbInformation, "Button Created"
End Sub
