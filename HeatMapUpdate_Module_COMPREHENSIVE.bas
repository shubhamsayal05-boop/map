Attribute VB_Name = "HeatMapUpdate_Comprehensive"
' ====================================================================
' Module: HeatMapUpdate_Comprehensive
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Date: 2025-11-23
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
    Dim statusCol As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    debugInfo = "=== HeatMap Status Update Debug Report ===" & vbCrLf & vbCrLf
    
    ' Step 1: Verify sheets exist
    debugInfo = debugInfo & "STEP 1: Verifying Sheets" & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetSheetList(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "✓ 'Evaluation Results' sheet found" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetSheetList(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugInfo = debugInfo & "✓ 'HeatMap Sheet' found" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing sheets..."
    
    ' Step 2: Analyze Evaluation Results sheet structure
    debugInfo = debugInfo & "STEP 2: Analyzing Evaluation Results Sheet" & vbCrLf
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "Total rows in Evaluation Results: " & lastRowEval & vbCrLf
    
    ' Find sections
    Dim overallStatusRow As Long
    Dim opModeSummaryRow As Long
    
    overallStatusRow = FindSectionRow(wsEval, "Overall Status by Op Code", lastRowEval)
    opModeSummaryRow = FindSectionRow(wsEval, "Operation Mode Summary", lastRowEval)
    
    If overallStatusRow > 0 Then
        debugInfo = debugInfo & "✓ 'Overall Status by Op Code' found at row " & overallStatusRow & vbCrLf
    Else
        debugInfo = debugInfo & "✗ 'Overall Status by Op Code' NOT found" & vbCrLf
    End If
    
    If opModeSummaryRow > 0 Then
        debugInfo = debugInfo & "✓ 'Operation Mode Summary' found at row " & opModeSummaryRow & vbCrLf
    Else
        debugInfo = debugInfo & "✗ 'Operation Mode Summary' NOT found" & vbCrLf
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Step 3: Analyze HeatMap Sheet structure
    debugInfo = debugInfo & "STEP 3: Analyzing HeatMap Sheet" & vbCrLf
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugInfo = debugInfo & "Total rows in HeatMap Sheet: " & lastRowHeatMap & vbCrLf
    
    ' Find Status column in HeatMap
    statusCol = FindStatusColumn(wsHeatMap)
    If statusCol > 0 Then
        debugInfo = debugInfo & "✓ 'Status' column found at column " & statusCol & " (" & ColumnLetter(statusCol) & ")" & vbCrLf
    Else
        debugInfo = debugInfo & "✗ 'Status' column NOT found - will use column C as default" & vbCrLf
        statusCol = 3 ' Default to column C
    End If
    
    ' Sample first 5 Op Codes from HeatMap
    debugInfo = debugInfo & vbCrLf & "Sample Op Codes in HeatMap (first 5):" & vbCrLf
    For i = 2 To Application.Min(6, lastRowHeatMap)
        If Trim(CStr(wsHeatMap.Cells(i, 1).Value)) <> "" Then
            debugInfo = debugInfo & "  Row " & i & ": " & Trim(CStr(wsHeatMap.Cells(i, 1).Value)) & vbCrLf
        End If
    Next i
    debugInfo = debugInfo & vbCrLf
    
    ' Step 4: Process Overall Status by Op Code section
    If overallStatusRow > 0 Then
        debugInfo = debugInfo & "STEP 4: Processing 'Overall Status by Op Code' Section" & vbCrLf
        Application.StatusBar = "Processing sub-operations..."
        
        Dim statusColEval As Long
        statusColEval = FindColumnInRow(wsEval, overallStatusRow + 1, "Final Status")
        
        If statusColEval > 0 Then
            debugInfo = debugInfo & "✓ 'Final Status' column found at column " & statusColEval & vbCrLf
            
            Dim processedCount As Long
            Dim matchedCount As Long
            processedCount = 0
            matchedCount = 0
            
            ' Process rows in this section
            For i = overallStatusRow + 2 To lastRowEval
                ' Stop if we hit next section or empty rows
                If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
                    Exit For
                End If
                
                If IsEmpty(wsEval.Cells(i, 1)) Or Trim(CStr(wsEval.Cells(i, 1).Value)) = "" Then
                    Exit For
                End If
                
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                    processedCount = processedCount + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColEval).Value)))
                    
                    ' Update HeatMap
                    If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap) Then
                        matchedCount = matchedCount + 1
                        updatedCount = updatedCount + 1
                    End If
                End If
            Next i
            
            debugInfo = debugInfo & "Processed " & processedCount & " operation codes" & vbCrLf
            debugInfo = debugInfo & "Matched and updated " & matchedCount & " operations" & vbCrLf
        Else
            debugInfo = debugInfo & "✗ 'Final Status' column NOT found in this section" & vbCrLf
        End If
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Step 5: Process Operation Mode Summary section
    If opModeSummaryRow > 0 Then
        debugInfo = debugInfo & "STEP 5: Processing 'Operation Mode Summary' Section" & vbCrLf
        Application.StatusBar = "Processing parent operations..."
        
        Dim statusColSummary As Long
        statusColSummary = FindColumnInRow(wsEval, opModeSummaryRow + 1, "Final Status")
        
        If statusColSummary > 0 Then
            debugInfo = debugInfo & "✓ 'Final Status' column found at column " & statusColSummary & vbCrLf
            
            Dim processedSummary As Long
            Dim matchedSummary As Long
            processedSummary = 0
            matchedSummary = 0
            
            ' Process rows in this section
            For i = opModeSummaryRow + 2 To lastRowEval
                ' Stop if we hit empty rows
                If IsEmpty(wsEval.Cells(i, 1)) Or Trim(CStr(wsEval.Cells(i, 1).Value)) = "" Then
                    Exit For
                End If
                
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                    processedSummary = processedSummary + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColSummary).Value)))
                    
                    ' Update HeatMap
                    If UpdateHeatMapRow(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap) Then
                        matchedSummary = matchedSummary + 1
                        updatedCount = updatedCount + 1
                    End If
                End If
            Next i
            
            debugInfo = debugInfo & "Processed " & processedSummary & " operation codes" & vbCrLf
            debugInfo = debugInfo & "Matched and updated " & matchedSummary & " operations" & vbCrLf
        Else
            debugInfo = debugInfo & "✗ 'Final Status' column NOT found in this section" & vbCrLf
        End If
    End If
    
    debugInfo = debugInfo & vbCrLf
    
    ' Completion
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim elapsed As Double
    elapsed = Round(Timer - startTime, 2)
    
    debugInfo = debugInfo & "=== SUMMARY ===" & vbCrLf
    debugInfo = debugInfo & "Total operations updated: " & updatedCount & vbCrLf
    debugInfo = debugInfo & "Time elapsed: " & elapsed & " seconds" & vbCrLf
    
    ' Show results
    If updatedCount > 0 Then
        MsgBox "HeatMap Status Update Complete!" & vbCrLf & vbCrLf & _
               "Updated " & updatedCount & " operations in " & elapsed & " seconds" & vbCrLf & vbCrLf & _
               "Click OK to see detailed debug report.", vbInformation, "Update Complete"
        
        ' Show debug info in message box (for detailed troubleshooting)
        MsgBox debugInfo, vbInformation, "Debug Report"
    Else
        MsgBox "WARNING: No operations were updated!" & vbCrLf & vbCrLf & _
               "Please review the debug report for details.", vbExclamation, "No Updates"
        
        ' Show debug info
        MsgBox debugInfo, vbInformation, "Debug Report"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error in UpdateHeatMapStatus:" & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, vbCritical, "Error"
End Sub

' Helper function to find a section row by title
Private Function FindSectionRow(ws As Worksheet, sectionTitle As String, lastRow As Long) As Long
    Dim i As Long
    Dim cellValue As String
    
    FindSectionRow = 0
    
    For i = 1 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 1).Value))
        If InStr(1, cellValue, sectionTitle, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find a column by header name in a specific row
Private Function FindColumnInRow(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindColumnInRow = 0
    
    For col = 1 To 50 ' Search first 50 columns
        cellValue = Trim(UCase(CStr(ws.Cells(headerRow, col).Value)))
        If InStr(1, cellValue, UCase(headerName), vbTextCompare) > 0 Then
            FindColumnInRow = col
            Exit Function
        End If
    Next col
End Function

' Helper function to find Status column in HeatMap sheet
Private Function FindStatusColumn(ws As Worksheet) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindStatusColumn = 0
    
    ' Search in first row for "Status" or "Current Status"
    For col = 1 To 30
        cellValue = Trim(UCase(CStr(ws.Cells(1, col).Value)))
        If InStr(1, cellValue, "STATUS", vbTextCompare) > 0 And _
           InStr(1, cellValue, "OPERATION", vbTextCompare) = 0 Then
            FindStatusColumn = col
            Exit Function
        End If
    Next col
End Function

' Helper function to update a row in HeatMap sheet
Private Function UpdateHeatMapRow(ws As Worksheet, opCode As String, status As String, statusCol As Long, lastRow As Long) As Boolean
    Dim i As Long
    Dim heatMapOpCode As String
    Dim statusDot As String
    Dim dotColor As Long
    
    UpdateHeatMapRow = False
    
    ' Skip if status is N/A or empty
    If status = "" Or status = "N/A" Then
        Exit Function
    End If
    
    ' Find matching operation code in HeatMap
    For i = 2 To lastRow
        heatMapOpCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If heatMapOpCode = opCode Then
            ' Determine dot character and color
            Select Case status
                Case "RED"
                    statusDot = "●"
                    dotColor = RGB(255, 0, 0) ' Red
                Case "YELLOW"
                    statusDot = "●"
                    dotColor = RGB(255, 192, 0) ' Yellow/Orange
                Case "GREEN"
                    statusDot = "●"
                    dotColor = RGB(0, 176, 80) ' Green
                Case Else
                    statusDot = "●"
                    dotColor = RGB(128, 128, 128) ' Gray
            End Select
            
            ' Update cell with colored dot
            With ws.Cells(i, statusCol)
                .Value = statusDot
                .Font.Name = "Wingdings"
                .Font.Size = 14
                .Font.Color = dotColor
                .HorizontalAlignment = xlCenter
                .VerticalAlignment = xlCenter
            End With
            
            UpdateHeatMapRow = True
            Exit Function
        End If
    Next i
End Function

' Helper function to get column letter from number
Private Function ColumnLetter(colNum As Long) As String
    ColumnLetter = Split(Cells(1, colNum).Address, "$")(1)
End Function

' Helper function to get list of all sheets
Private Function GetSheetList() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        sheetList = sheetList & "  - " & ws.Name & vbCrLf
    Next ws
    
    GetSheetList = sheetList
End Function

' Function to create Update button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim wsHeatMap As Worksheet
    Dim btn As Button
    Dim btnRange As Range
    
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetSheetList(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Delete existing button if present
    On Error Resume Next
    wsHeatMap.Buttons("btnUpdateHeatMap").Delete
    On Error GoTo 0
    
    ' Create button in cell B2:D3
    Set btnRange = wsHeatMap.Range("B2:D3")
    Set btn = wsHeatMap.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    
    With btn
        .Name = "btnUpdateHeatMap"
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the 'Update HeatMap Status' button to transfer evaluation results.", _
           vbInformation, "Button Created"
End Sub
