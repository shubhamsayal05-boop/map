Attribute VB_Name = "HeatMapUpdate_Debug"
' ====================================================================
' Module: HeatMapUpdate_Debug
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Version: 2.0 - Improved with detailed diagnostics
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
    Dim debugReport As String
    Dim evalOpCodes As Long, heatMapOpCodes As Long
    Dim matchedCodes As Long, skippedNA As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalOpCodes = 0
    heatMapOpCodes = 0
    matchedCodes = 0
    skippedNA = 0
    debugReport = "=== HEATMAP UPDATE DEBUG REPORT ===" & vbCrLf & vbCrLf
    
    ' Step 1: Get worksheets
    debugReport = debugReport & "STEP 1: Locating Worksheets" & vbCrLf
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetList(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugReport = debugReport & "  ✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets: " & GetSheetList(), vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugReport = debugReport & "  ✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing data structure..."
    
    ' Step 2: Analyze sheet structure
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugReport = debugReport & "STEP 2: Sheet Analysis" & vbCrLf
    debugReport = debugReport & "  Evaluation Results: " & lastRowEval & " total rows" & vbCrLf
    debugReport = debugReport & "  HeatMap Sheet: " & lastRowHeatMap & " total rows" & vbCrLf & vbCrLf
    
    ' Step 3: Find "Overall Status by Op Code" section
    debugReport = debugReport & "STEP 3: Locating Data Sections" & vbCrLf
    Dim overallStartRow As Long
    overallStartRow = FindSectionRow(wsEval, "Overall Status by Op Code", lastRowEval)
    
    If overallStartRow > 0 Then
        debugReport = debugReport & "  ✓ 'Overall Status by Op Code' found at row " & overallStartRow & vbCrLf
        
        ' Find the header row and Final Status column
        Dim headerRow As Long
        headerRow = overallStartRow + 1
        debugReport = debugReport & "  Header row: " & headerRow & vbCrLf
        debugReport = debugReport & "  Headers: " & GetRowContent(wsEval, headerRow, 5) & vbCrLf
        
        Dim statusCol As Long
        statusCol = FindColumnByHeader(wsEval, headerRow, "Final Status")
        If statusCol > 0 Then
            debugReport = debugReport & "  ✓ 'Final Status' column found at: " & Chr(64 + statusCol) & " (column " & statusCol & ")" & vbCrLf
        Else
            debugReport = debugReport & "  ✗ 'Final Status' column NOT found in header row" & vbCrLf
            statusCol = FindColumnByHeader(wsEval, headerRow, "Overall Status")
            If statusCol > 0 Then
                debugReport = debugReport & "  ✓ Using 'Overall Status' column instead at: " & Chr(64 + statusCol) & " (column " & statusCol & ")" & vbCrLf
            End If
        End If
        
        debugReport = debugReport & vbCrLf & "STEP 4: Processing Overall Status Section" & vbCrLf
        
        ' Process Overall Status by Op Code section
        For i = headerRow + 1 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Stop if we hit the next section
            If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
                debugReport = debugReport & "  End of section at row " & i & vbCrLf
                Exit For
            End If
            
            ' Check if this is a valid operation code
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 7 Then
                evalOpCodes = evalOpCodes + 1
                
                ' Get Final Status
                If statusCol > 0 Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusCol).Value)))
                Else
                    finalStatus = ""
                End If
                
                ' Debug: Show first few operations
                If evalOpCodes <= 3 Then
                    debugReport = debugReport & "  Sample " & evalOpCodes & ": Code=" & opCode & _
                                  ", Status=" & finalStatus & " (row " & i & ")" & vbCrLf
                End If
                
                ' Skip N/A statuses
                If finalStatus = "N/A" Or finalStatus = "" Then
                    skippedNA = skippedNA + 1
                ElseIf finalStatus = "RED" Or finalStatus = "YELLOW" Or finalStatus = "GREEN" Then
                    ' Try to update in HeatMap
                    If UpdateOperationStatusDebug(wsHeatMap, opCode, finalStatus, lastRowHeatMap, debugReport) Then
                        updatedCount = updatedCount + 1
                        matchedCodes = matchedCodes + 1
                    End If
                End If
            End If
        Next i
        
        debugReport = debugReport & "  Total operation codes found: " & evalOpCodes & vbCrLf
        debugReport = debugReport & "  Skipped (N/A or blank): " & skippedNA & vbCrLf
        debugReport = debugReport & "  Matched and updated: " & matchedCodes & vbCrLf & vbCrLf
    Else
        debugReport = debugReport & "  ✗ 'Overall Status by Op Code' section NOT found!" & vbCrLf & vbCrLf
    End If
    
    ' Step 5: Find and process "Operation Mode Summary" section
    debugReport = debugReport & "STEP 5: Processing Operation Mode Summary Section" & vbCrLf
    Dim summaryStartRow As Long
    summaryStartRow = FindSectionRow(wsEval, "Operation Mode Summary", lastRowEval)
    
    If summaryStartRow > 0 Then
        debugReport = debugReport & "  ✓ 'Operation Mode Summary' found at row " & summaryStartRow & vbCrLf
        
        ' Find the header row and columns
        Dim summaryHeaderRow As Long
        summaryHeaderRow = summaryStartRow + 1
        debugReport = debugReport & "  Header row: " & summaryHeaderRow & vbCrLf
        debugReport = debugReport & "  Headers: " & GetRowContent(wsEval, summaryHeaderRow, 10) & vbCrLf
        
        Dim opCodeColSummary As Long, statusColSummary As Long
        opCodeColSummary = FindColumnByHeader(wsEval, summaryHeaderRow, "Op Code")
        statusColSummary = FindColumnByHeader(wsEval, summaryHeaderRow, "Final Status")
        
        If opCodeColSummary = 0 Then
            opCodeColSummary = 6 ' Try column F as default
            debugReport = debugReport & "  Using default Op Code column: F (column " & opCodeColSummary & ")" & vbCrLf
        Else
            debugReport = debugReport & "  ✓ Op Code column: " & Chr(64 + opCodeColSummary) & " (column " & opCodeColSummary & ")" & vbCrLf
        End If
        
        If statusColSummary = 0 Then
            statusColSummary = 9 ' Try column I as default
            debugReport = debugReport & "  Using default Final Status column: I (column " & statusColSummary & ")" & vbCrLf
        Else
            debugReport = debugReport & "  ✓ Final Status column: " & Chr(64 + statusColSummary) & " (column " & statusColSummary & ")" & vbCrLf
        End If
        
        Dim summaryCount As Long
        summaryCount = 0
        
        ' Process Operation Mode Summary section
        For i = summaryHeaderRow + 1 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, opCodeColSummary).Value))
            
            ' Stop at empty rows or next section
            If opCode = "" Then
                Dim nextContent As String
                nextContent = Trim(CStr(wsEval.Cells(i, 1).Value))
                If nextContent <> "" Then Exit For ' Hit another section
            End If
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 7 Then
                summaryCount = summaryCount + 1
                
                ' Get Final Status
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColSummary).Value)))
                
                ' Debug: Show first few operations
                If summaryCount <= 3 Then
                    debugReport = debugReport & "  Sample " & summaryCount & ": Code=" & opCode & _
                                  ", Status=" & finalStatus & " (row " & i & ")" & vbCrLf
                End If
                
                ' Update in HeatMap
                If finalStatus <> "N/A" And finalStatus <> "" Then
                    If finalStatus = "RED" Or finalStatus = "YELLOW" Or finalStatus = "GREEN" Then
                        If UpdateOperationStatusDebug(wsHeatMap, opCode, finalStatus, lastRowHeatMap, debugReport) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            End If
        Next i
        
        debugReport = debugReport & "  Parent operation codes found: " & summaryCount & vbCrLf & vbCrLf
    Else
        debugReport = debugReport & "  ✗ 'Operation Mode Summary' section NOT found!" & vbCrLf & vbCrLf
    End If
    
    ' Step 6: Analyze HeatMap Sheet structure
    debugReport = debugReport & "STEP 6: HeatMap Sheet Analysis" & vbCrLf
    Dim heatMapStatusCol As Long
    heatMapStatusCol = FindColumnByHeader(wsHeatMap, 1, "Status")
    If heatMapStatusCol = 0 Then
        ' Try to find by looking for common status column headers
        For j = 1 To 10
            Dim headerVal As String
            headerVal = Trim(UCase(CStr(wsHeatMap.Cells(1, j).Value)))
            If InStr(1, headerVal, "STATUS", vbTextCompare) > 0 Or _
               InStr(1, headerVal, "CURRENT", vbTextCompare) > 0 Then
                heatMapStatusCol = j
                Exit For
            End If
        Next j
    End If
    
    If heatMapStatusCol > 0 Then
        debugReport = debugReport & "  ✓ Status column found at: " & Chr(64 + heatMapStatusCol) & " (column " & heatMapStatusCol & ")" & vbCrLf
    Else
        debugReport = debugReport & "  Using default status column: C (column 3)" & vbCrLf
        heatMapStatusCol = 3
    End If
    
    ' Count operation codes in HeatMap
    For i = 2 To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" And IsNumeric(opCode) And Len(opCode) >= 7 Then
            heatMapOpCodes = heatMapOpCodes + 1
        End If
    Next i
    debugReport = debugReport & "  Total operation codes in HeatMap: " & heatMapOpCodes & vbCrLf & vbCrLf
    
    ' Final Summary
    Dim duration As Double
    duration = Round(Timer - startTime, 2)
    
    debugReport = debugReport & "=== FINAL SUMMARY ===" & vbCrLf
    debugReport = debugReport & "  Operations updated: " & updatedCount & vbCrLf
    debugReport = debugReport & "  Time elapsed: " & duration & " seconds" & vbCrLf & vbCrLf
    
    If updatedCount = 0 Then
        debugReport = debugReport & "⚠️ NO OPERATIONS WERE UPDATED!" & vbCrLf & vbCrLf
        debugReport = debugReport & "Possible Issues:" & vbCrLf
        debugReport = debugReport & "  1. Operation codes don't match between sheets" & vbCrLf
        debugReport = debugReport & "  2. All statuses are N/A" & vbCrLf
        debugReport = debugReport & "  3. Column headers are different than expected" & vbCrLf
        debugReport = debugReport & "  4. Data sections not found correctly" & vbCrLf
    End If
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    ' Show the debug report
    MsgBox debugReport, vbInformation, "HeatMap Update Debug Report"
    
    ' Option to view in a text window for easier reading
    If updatedCount = 0 Then
        If MsgBox("Would you like to copy the full debug report to clipboard?", vbQuestion + vbYesNo, "Copy Debug Info") = vbYes Then
            CreateTextFile debugReport
        End If
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Debug Info:" & vbCrLf & debugReport, vbCritical, "Update Error"
End Sub

' Helper function to find a section by name
Private Function FindSectionRow(ws As Worksheet, sectionName As String, lastRow As Long) As Long
    Dim i As Long
    FindSectionRow = 0
    
    For i = 1 To lastRow
        If InStr(1, CStr(ws.Cells(i, 1).Value), sectionName, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
    Next i
End Function

' Helper function to find a column by header name
Private Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim j As Long
    Dim headerVal As String
    FindColumnByHeader = 0
    
    For j = 1 To 30 ' Check first 30 columns
        headerVal = Trim(UCase(CStr(ws.Cells(headerRow, j).Value)))
        If InStr(1, headerVal, UCase(headerName), vbTextCompare) > 0 Then
            FindColumnByHeader = j
            Exit Function
        End If
    Next j
End Function

' Helper function to update operation status in HeatMap with debugging
Private Function UpdateOperationStatusDebug(ws As Worksheet, opCode As String, status As String, _
                                            lastRow As Long, ByRef debugMsg As String) As Boolean
    Dim i As Long
    Dim heatMapCode As String
    Dim statusCol As Long
    UpdateOperationStatusDebug = False
    
    ' Find status column (default to column 3 if not found)
    statusCol = FindColumnByHeader(ws, 1, "Status")
    If statusCol = 0 Then statusCol = 3
    
    ' Search for the operation code in HeatMap
    For i = 2 To lastRow
        heatMapCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If heatMapCode = opCode Then
            ' Found matching code - update status with colored dot
            Dim dotChar As String
            Dim dotColor As Long
            
            dotChar = "●" ' Filled circle character
            
            Select Case UCase(status)
                Case "RED"
                    dotColor = RGB(255, 0, 0)     ' Red
                Case "YELLOW"
                    dotColor = RGB(255, 255, 0)   ' Yellow
                Case "GREEN"
                    dotColor = RGB(0, 255, 0)     ' Green
                Case Else
                    dotColor = RGB(128, 128, 128) ' Gray for N/A
                    dotChar = "○"                  ' Empty circle
            End Select
            
            ' Update the cell
            ws.Cells(i, statusCol).Value = dotChar
            ws.Cells(i, statusCol).Font.Color = dotColor
            ws.Cells(i, statusCol).Font.Size = 14
            ws.Cells(i, statusCol).Font.Name = "Arial"
            
            UpdateOperationStatusDebug = True
            Exit Function
        End If
    Next i
    
    ' If we get here, the operation code wasn't found in HeatMap
End Function

' Helper function to get row content for debugging
Private Function GetRowContent(ws As Worksheet, rowNum As Long, numCols As Long) As String
    Dim j As Long
    Dim content As String
    content = ""
    
    For j = 1 To numCols
        Dim cellVal As String
        cellVal = Trim(CStr(ws.Cells(rowNum, j).Value))
        If cellVal <> "" Then
            If content <> "" Then content = content & " | "
            content = content & Chr(64 + j) & ":" & cellVal
        End If
    Next j
    
    GetRowContent = content
End Function

' Helper function to get list of all sheet names
Private Function GetSheetList() As String
    Dim ws As Worksheet
    Dim sheetList As String
    sheetList = ""
    
    For Each ws In ThisWorkbook.Worksheets
        If sheetList <> "" Then sheetList = sheetList & ", "
        sheetList = sheetList & ws.Name
    Next ws
    
    GetSheetList = sheetList
End Function

' Helper function to create a text file with debug info
Private Sub CreateTextFile(content As String)
    Dim objData As DataObject
    Set objData = New DataObject
    objData.SetText content
    objData.PutInClipboard
    MsgBox "Debug report copied to clipboard!" & vbCrLf & vbCrLf & _
           "You can paste it into Notepad or any text editor.", vbInformation
End Sub

' Function to create the update button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnRange As Range
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if it exists
    Dim shp As Shape
    For Each shp In ws.Shapes
        If shp.Name = "UpdateHeatMapButton" Or shp.Type = msoFormControl Then
            If InStr(1, shp.Name, "Button", vbTextCompare) > 0 Then
                shp.Delete
            End If
        End If
    Next shp
    
    ' Create button in cell B2
    Set btnRange = ws.Range("B2:D3")
    Set btn = ws.Buttons.Add(btnRange.Left, btnRange.Top, btnRange.Width, btnRange.Height)
    
    With btn
        .Name = "UpdateHeatMapButton"
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update operation statuses after running evaluation.", _
           vbInformation, "Button Created"
End Sub
