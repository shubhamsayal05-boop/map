Attribute VB_Name = "HeatMapUpdate_Module_COMPLETE_DEBUG"
' ===================================================================
' HeatMap Status Update Module - Complete with Enhanced Debugging
' ===================================================================
' This module transfers evaluation results to HeatMap Sheet with
' comprehensive error messages and diagnostic information
'
' Version: Complete Debug v1.0
' Last Updated: 2025-11-23
' ===================================================================

Option Explicit

' Module-level variables for tracking
Private m_UpdatedCount As Long
Private m_TotalProcessed As Long
Private m_ErrorsFound As Collection

' ===================================================================
' Main Update Function with Enhanced Debugging
' ===================================================================
Sub UpdateHeatMapStatus()
    On Error GoTo ErrorHandler
    
    Dim wsEvalResults As Worksheet
    Dim wsHeatMap As Worksheet
    Dim lastRow As Long
    Dim i As Long
    Dim opCode As String
    Dim finalStatus As String
    Dim heatMapRow As Long
    Dim startTime As Double
    Dim debugMsg As String
    Dim foundEvalSheet As Boolean
    Dim foundHeatMapSheet As Boolean
    
    ' Initialize tracking
    m_UpdatedCount = 0
    m_TotalProcessed = 0
    Set m_ErrorsFound = New Collection
    startTime = Timer
    
    ' ===============================================================
    ' STEP 1: Validate Workbook and Sheets
    ' ===============================================================
    debugMsg = "DIAGNOSTIC REPORT" & vbCrLf & String(60, "=") & vbCrLf & vbCrLf
    
    ' List all available sheets
    debugMsg = debugMsg & "Available Sheets in Workbook:" & vbCrLf
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        debugMsg = debugMsg & "  - " & ws.Name & vbCrLf
    Next ws
    debugMsg = debugMsg & vbCrLf
    
    ' Try to find Evaluation Results sheet
    foundEvalSheet = False
    On Error Resume Next
    Set wsEvalResults = ThisWorkbook.Sheets("Evaluation Results")
    If Not wsEvalResults Is Nothing Then
        foundEvalSheet = True
        debugMsg = debugMsg & "✓ Found 'Evaluation Results' sheet" & vbCrLf
    Else
        debugMsg = debugMsg & "✗ 'Evaluation Results' sheet NOT FOUND" & vbCrLf
        debugMsg = debugMsg & "  Looking for alternative names..." & vbCrLf
        
        ' Try alternative names
        Set wsEvalResults = ThisWorkbook.Sheets("Evaluation_Results")
        If Not wsEvalResults Is Nothing Then
            foundEvalSheet = True
            debugMsg = debugMsg & "✓ Found 'Evaluation_Results' sheet (underscore)" & vbCrLf
        Else
            Set wsEvalResults = ThisWorkbook.Sheets("EvaluationResults")
            If Not wsEvalResults Is Nothing Then
                foundEvalSheet = True
                debugMsg = debugMsg & "✓ Found 'EvaluationResults' sheet (no space)" & vbCrLf
            End If
        End If
    End If
    On Error GoTo ErrorHandler
    
    If Not foundEvalSheet Then
        MsgBox debugMsg & vbCrLf & "ERROR: Cannot find Evaluation Results sheet!" & vbCrLf & _
               "Please ensure you have run the evaluation first.", vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Try to find HeatMap sheet
    foundHeatMapSheet = False
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If Not wsHeatMap Is Nothing Then
        foundHeatMapSheet = True
        debugMsg = debugMsg & "✓ Found 'HeatMap Sheet'" & vbCrLf
    Else
        debugMsg = debugMsg & "✗ 'HeatMap Sheet' NOT FOUND" & vbCrLf
        debugMsg = debugMsg & "  Looking for alternative names..." & vbCrLf
        
        ' Try alternative names
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap")
        If Not wsHeatMap Is Nothing Then
            foundHeatMapSheet = True
            debugMsg = debugMsg & "✓ Found 'HeatMap' sheet" & vbCrLf
        Else
            Set wsHeatMap = ThisWorkbook.Sheets("Heat Map")
            If Not wsHeatMap Is Nothing Then
                foundHeatMapSheet = True
                debugMsg = debugMsg & "✓ Found 'Heat Map' sheet (with space)" & vbCrLf
            Else
                Set wsHeatMap = ThisWorkbook.Sheets("HeatMap_Template")
                If Not wsHeatMap Is Nothing Then
                    foundHeatMapSheet = True
                    debugMsg = debugMsg & "✓ Found 'HeatMap_Template' sheet" & vbCrLf
                End If
            End If
        End If
    End If
    On Error GoTo ErrorHandler
    
    If Not foundHeatMapSheet Then
        MsgBox debugMsg & vbCrLf & "ERROR: Cannot find HeatMap Sheet!" & vbCrLf & _
               "Please ensure the sheet exists.", vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' ===============================================================
    ' STEP 2: Analyze Evaluation Results Structure
    ' ===============================================================
    debugMsg = debugMsg & "Evaluation Results Sheet Analysis:" & vbCrLf
    debugMsg = debugMsg & "  Sheet Name: " & wsEvalResults.Name & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    Dim overallStatusRow As Long
    Dim opModeSummaryRow As Long
    overallStatusRow = 0
    opModeSummaryRow = 0
    
    ' Search for section headers
    For i = 1 To 100
        Dim cellVal As String
        cellVal = Trim(wsEvalResults.Cells(i, 1).Value)
        
        If InStr(1, cellVal, "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallStatusRow = i
            debugMsg = debugMsg & "  ✓ Found 'Overall Status by Op Code' at row " & i & vbCrLf
        End If
        
        If InStr(1, cellVal, "Operation Mode Summary", vbTextCompare) > 0 Then
            opModeSummaryRow = i
            debugMsg = debugMsg & "  ✓ Found 'Operation Mode Summary' at row " & i & vbCrLf
        End If
        
        If overallStatusRow > 0 And opModeSummaryRow > 0 Then Exit For
    Next i
    
    If overallStatusRow = 0 Then
        debugMsg = debugMsg & "  ✗ 'Overall Status by Op Code' section NOT FOUND" & vbCrLf
    End If
    
    If opModeSummaryRow = 0 Then
        debugMsg = debugMsg & "  ✗ 'Operation Mode Summary' section NOT FOUND" & vbCrLf
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' Find last row with data in Evaluation Results
    lastRow = wsEvalResults.Cells(wsEvalResults.Rows.Count, 1).End(xlUp).Row
    debugMsg = debugMsg & "  Last row with data: " & lastRow & vbCrLf
    
    ' Sample first few operation codes
    debugMsg = debugMsg & "  Sample operation codes from Evaluation Results:" & vbCrLf
    Dim sampleCount As Long
    sampleCount = 0
    For i = overallStatusRow + 2 To lastRow
        If sampleCount >= 5 Then Exit For
        opCode = Trim(wsEvalResults.Cells(i, 1).Value)
        If opCode <> "" And IsNumeric(opCode) Then
            finalStatus = Trim(wsEvalResults.Cells(i, 3).Value) ' Column C typically has status
            debugMsg = debugMsg & "    Row " & i & ": OpCode=" & opCode & ", Status=" & finalStatus & vbCrLf
            sampleCount = sampleCount + 1
        End If
    Next i
    
    debugMsg = debugMsg & vbCrLf
    
    ' ===============================================================
    ' STEP 3: Analyze HeatMap Sheet Structure
    ' ===============================================================
    debugMsg = debugMsg & "HeatMap Sheet Analysis:" & vbCrLf
    debugMsg = debugMsg & "  Sheet Name: " & wsHeatMap.Name & vbCrLf
    
    ' Find last row in HeatMap
    Dim heatMapLastRow As Long
    heatMapLastRow = wsHeatMap.Cells(wsHeatMap.Rows.Count, 1).End(xlUp).Row
    debugMsg = debugMsg & "  Last row with data: " & heatMapLastRow & vbCrLf
    
    ' Find Status column
    Dim statusCol As Long
    statusCol = 0
    For i = 1 To 20
        cellVal = Trim(UCase(wsHeatMap.Cells(1, i).Value))
        If InStr(cellVal, "STATUS") > 0 Or cellVal = "STATUS" Then
            statusCol = i
            debugMsg = debugMsg & "  ✓ Found 'Status' column at column " & i & " (" & Split(Cells(1, i).Address, "$")(1) & ")" & vbCrLf
            Exit For
        End If
    Next i
    
    If statusCol = 0 Then
        debugMsg = debugMsg & "  ✗ 'Status' column NOT FOUND in first 20 columns" & vbCrLf
        debugMsg = debugMsg & "  Will assume column 3 (C) for status" & vbCrLf
        statusCol = 3
    End If
    
    ' Sample first few operation codes from HeatMap
    debugMsg = debugMsg & "  Sample operation codes from HeatMap:" & vbCrLf
    sampleCount = 0
    For i = 2 To heatMapLastRow
        If sampleCount >= 5 Then Exit For
        opCode = Trim(wsHeatMap.Cells(i, 1).Value)
        If opCode <> "" And IsNumeric(opCode) Then
            debugMsg = debugMsg & "    Row " & i & ": OpCode=" & opCode & vbCrLf
            sampleCount = sampleCount + 1
        End If
    Next i
    
    debugMsg = debugMsg & vbCrLf
    
    ' ===============================================================
    ' STEP 4: Process and Update Statuses
    ' ===============================================================
    debugMsg = debugMsg & "Processing Updates:" & vbCrLf
    
    ' Process sub-operations from "Overall Status by Op Code" section
    If overallStatusRow > 0 Then
        For i = overallStatusRow + 2 To lastRow
            opCode = Trim(wsEvalResults.Cells(i, 1).Value)
            
            ' Stop if we reach another section
            If opCode = "" Or InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                Exit For
            End If
            
            If IsNumeric(opCode) And Len(opCode) = 8 Then
                finalStatus = Trim(wsEvalResults.Cells(i, 3).Value) ' Column C for status
                
                ' Find this opCode in HeatMap
                heatMapRow = FindOperationInHeatMap(wsHeatMap, opCode)
                
                If heatMapRow > 0 Then
                    ' Update status
                    Call SetStatusInHeatMap(wsHeatMap, heatMapRow, statusCol, finalStatus)
                    m_UpdatedCount = m_UpdatedCount + 1
                End If
                
                m_TotalProcessed = m_TotalProcessed + 1
            End If
        Next i
    End If
    
    ' Process parent operations from "Operation Mode Summary" section
    If opModeSummaryRow > 0 Then
        For i = opModeSummaryRow + 2 To lastRow
            opCode = Trim(wsEvalResults.Cells(i, 1).Value)
            
            If opCode = "" Then Exit For
            
            If IsNumeric(opCode) And Len(opCode) = 8 Then
                ' For Operation Mode Summary, status might be in different column
                ' Try column 3 first, then column 4
                finalStatus = Trim(wsEvalResults.Cells(i, 3).Value)
                If finalStatus = "" Or finalStatus = "N/A" Then
                    finalStatus = Trim(wsEvalResults.Cells(i, 4).Value)
                End If
                
                ' Find this opCode in HeatMap
                heatMapRow = FindOperationInHeatMap(wsHeatMap, opCode)
                
                If heatMapRow > 0 Then
                    ' Update status
                    Call SetStatusInHeatMap(wsHeatMap, heatMapRow, statusCol, finalStatus)
                    m_UpdatedCount = m_UpdatedCount + 1
                End If
                
                m_TotalProcessed = m_TotalProcessed + 1
            End If
        Next i
    End If
    
    ' ===============================================================
    ' STEP 5: Display Results
    ' ===============================================================
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    debugMsg = debugMsg & vbCrLf & String(60, "=") & vbCrLf
    debugMsg = debugMsg & "RESULTS:" & vbCrLf
    debugMsg = debugMsg & "  Operations processed: " & m_TotalProcessed & vbCrLf
    debugMsg = debugMsg & "  Statuses updated: " & m_UpdatedCount & vbCrLf
    debugMsg = debugMsg & "  Time elapsed: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf
    
    If m_UpdatedCount = 0 Then
        debugMsg = debugMsg & vbCrLf & "WARNING: No statuses were updated!" & vbCrLf
        debugMsg = debugMsg & "This usually means:" & vbCrLf
        debugMsg = debugMsg & "  1. Operation codes don't match between sheets" & vbCrLf
        debugMsg = debugMsg & "  2. Evaluation hasn't been run" & vbCrLf
        debugMsg = debugMsg & "  3. Sheet structure is different than expected" & vbCrLf
        
        MsgBox debugMsg, vbExclamation, "HeatMap Update - No Updates"
    Else
        MsgBox debugMsg, vbInformation, "HeatMap Update Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error in UpdateHeatMapStatus:" & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Description: " & Err.Description & vbCrLf & vbCrLf & _
           debugMsg, vbCritical, "Error"
End Sub

' ===================================================================
' Find Operation Code in HeatMap Sheet
' ===================================================================
Private Function FindOperationInHeatMap(ws As Worksheet, opCode As String) As Long
    Dim i As Long
    Dim lastRow As Long
    Dim cellValue As String
    
    FindOperationInHeatMap = 0
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    ' Search for exact match
    For i = 2 To lastRow
        cellValue = Trim(ws.Cells(i, 1).Value)
        If cellValue = opCode Then
            FindOperationInHeatMap = i
            Exit Function
        End If
    Next i
End Function

' ===================================================================
' Set Status in HeatMap with Color-Coded Dot
' ===================================================================
Private Sub SetStatusInHeatMap(ws As Worksheet, rowNum As Long, colNum As Long, status As String)
    Dim statusUpper As String
    Dim dotChar As String
    Dim colorRGB As Long
    
    statusUpper = UCase(Trim(status))
    dotChar = ChrW(9679) ' Filled circle Unicode character
    
    ' Determine color based on status
    Select Case statusUpper
        Case "RED"
            colorRGB = RGB(255, 0, 0) ' Red
        Case "YELLOW"
            colorRGB = RGB(255, 255, 0) ' Yellow
        Case "GREEN"
            colorRGB = RGB(0, 255, 0) ' Green
        Case Else
            colorRGB = RGB(128, 128, 128) ' Gray for N/A
    End Select
    
    ' Set the cell value and formatting
    With ws.Cells(rowNum, colNum)
        .Value = dotChar
        .Font.Name = "Wingdings"
        .Font.Size = 14
        .Font.Color = colorRGB
        .HorizontalAlignment = xlCenter
    End With
End Sub

' ===================================================================
' Create Update Button on HeatMap Sheet
' ===================================================================
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    
    ' Find HeatMap sheet
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("HeatMap")
    End If
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("Heat Map")
    End If
    
    On Error GoTo ErrorHandler
    
    If ws Is Nothing Then
        MsgBox "Cannot find HeatMap Sheet to create button." & vbCrLf & _
               "Please ensure the sheet exists.", vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Delete existing button if it exists
    btnName = "btnUpdateHeatMap"
    On Error Resume Next
    ws.Buttons(btnName).Delete
    On Error GoTo ErrorHandler
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 150, 30)
    btn.Name = btnName
    btn.Text = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    MsgBox "Button created successfully on " & ws.Name & " sheet!" & vbCrLf & vbCrLf & _
           "Click the 'Update HeatMap Status' button to transfer evaluation results.", _
           vbInformation, "Button Created"
    
    Exit Sub
    
ErrorHandler:
    MsgBox "Error creating button:" & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbCritical, "Error"
End Sub
