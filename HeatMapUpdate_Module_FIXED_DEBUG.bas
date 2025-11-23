Attribute VB_Name = "HeatMapUpdate_Debug"
' ====================================================================
' Module: HeatMapUpdate_Debug
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Version: 3.0 - Fixed for correct sheet structure
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
    Dim evalRowsFound As Long
    Dim heatMapRowsFound As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalRowsFound = 0
    heatMapRowsFound = 0
    debugMsg = ""
    
    ' === STEP 1: Verify Sheets Exist ===
    Application.StatusBar = "Step 1: Verifying sheets..."
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Sheet names found: " & ListAllSheets(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        ' Try alternative name
        Set wsHeatMap = ThisWorkbook.Sheets("HeatMap_Sheet")
        If wsHeatMap Is Nothing Then
            Set wsHeatMap = ThisWorkbook.Sheets("Heatmap Sheet")
            If wsHeatMap Is Nothing Then
                MsgBox "ERROR: HeatMap sheet not found!" & vbCrLf & vbCrLf & _
                       "Tried: 'HeatMap Sheet', 'HeatMap_Sheet', 'Heatmap Sheet'" & vbCrLf & _
                       "Sheet names found: " & ListAllSheets(), _
                       vbCritical, "Sheet Not Found"
                Exit Sub
            End If
        End If
    End If
    On Error GoTo ErrorHandler
    
    debugMsg = "✓ Sheets Found:" & vbCrLf & _
               "  - Evaluation Results: " & wsEval.Name & vbCrLf & _
               "  - HeatMap: " & wsHeatMap.Name & vbCrLf & vbCrLf
    
    ' Disable screen updating for performance
    Application.ScreenUpdating = False
    
    ' === STEP 2: Analyze Evaluation Results Structure ===
    Application.StatusBar = "Step 2: Analyzing Evaluation Results sheet..."
    
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "Evaluation Results Analysis:" & vbCrLf & _
               "  - Total rows: " & lastRowEval & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    Dim overallRow As Long
    overallRow = FindSectionRow(wsEval, "Overall Status by Op Code", lastRowEval)
    
    If overallRow > 0 Then
        debugMsg = debugMsg & "  - 'Overall Status by Op Code' found at row " & overallRow & vbCrLf
    Else
        debugMsg = debugMsg & "  - WARNING: 'Overall Status by Op Code' NOT FOUND!" & vbCrLf
    End If
    
    ' Find "Operation Mode Summary" section
    Dim summaryRow As Long
    summaryRow = FindSectionRow(wsEval, "Operation Mode Summary", lastRowEval)
    
    If summaryRow > 0 Then
        debugMsg = debugMsg & "  - 'Operation Mode Summary' found at row " & summaryRow & vbCrLf
    Else
        debugMsg = debugMsg & "  - WARNING: 'Operation Mode Summary' NOT FOUND!" & vbCrLf
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' === STEP 3: Analyze HeatMap Sheet Structure ===
    Application.StatusBar = "Step 3: Analyzing HeatMap Sheet..."
    
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "HeatMap Sheet Analysis:" & vbCrLf & _
               "  - Total rows: " & lastRowHeatMap & vbCrLf
    
    ' Find Status column in HeatMap
    Dim statusColHeatMap As Long
    statusColHeatMap = FindColumnByHeader(wsHeatMap, 1, "Status")
    
    If statusColHeatMap > 0 Then
        debugMsg = debugMsg & "  - 'Status' column found at column " & ColumnLetter(statusColHeatMap) & " (" & statusColHeatMap & ")" & vbCrLf
    Else
        ' Try alternative column names
        statusColHeatMap = FindColumnByHeader(wsHeatMap, 1, "Current Status")
        If statusColHeatMap = 0 Then
            statusColHeatMap = FindColumnByHeader(wsHeatMap, 1, "Current Status P1")
        End If
        
        If statusColHeatMap > 0 Then
            debugMsg = debugMsg & "  - Status column found at column " & ColumnLetter(statusColHeatMap) & " (" & statusColHeatMap & ")" & vbCrLf
        Else
            debugMsg = debugMsg & "  - WARNING: Status column NOT FOUND!" & vbCrLf & _
                       "    Tried: 'Status', 'Current Status', 'Current Status P1'" & vbCrLf
        End If
    End If
    
    ' Count operation codes in HeatMap
    For i = 2 To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" And IsNumeric(opCode) Then
            heatMapRowsFound = heatMapRowsFound + 1
        End If
    Next i
    
    debugMsg = debugMsg & "  - Operation codes found: " & heatMapRowsFound & vbCrLf & vbCrLf
    
    ' === STEP 4: Process Overall Status Section ===
    If overallRow > 0 Then
        Application.StatusBar = "Step 4: Processing Overall Status section..."
        
        Dim statusColOverall As Long
        ' Find "Final Status" or "Overall Status" column
        statusColOverall = FindColumnByHeader(wsEval, overallRow + 1, "Final Status")
        If statusColOverall = 0 Then
            statusColOverall = FindColumnByHeader(wsEval, overallRow + 1, "Overall Status")
        End If
        
        If statusColOverall > 0 Then
            debugMsg = debugMsg & "Processing Overall Status Section:" & vbCrLf & _
                       "  - Status column: " & ColumnLetter(statusColOverall) & " (" & statusColOverall & ")" & vbCrLf
            
            ' Process each operation in this section
            For i = overallRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
                
                ' Stop if we hit next section or empty
                If opCode = "" Or InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
                    Exit For
                End If
                
                If IsNumeric(opCode) Then
                    evalRowsFound = evalRowsFound + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColOverall).Value)))
                    
                    ' Update HeatMap if valid status
                    If finalStatus <> "" And finalStatus <> "N/A" And statusColHeatMap > 0 Then
                        If UpdateOperationInHeatMap(wsHeatMap, opCode, finalStatus, statusColHeatMap, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
            
            debugMsg = debugMsg & "  - Operations found: " & evalRowsFound & vbCrLf & _
                       "  - Operations updated: " & updatedCount & vbCrLf & vbCrLf
        Else
            debugMsg = debugMsg & "  - ERROR: Could not find status column in Overall Status section!" & vbCrLf & vbCrLf
        End If
    End If
    
    ' === STEP 5: Process Operation Mode Summary Section ===
    If summaryRow > 0 Then
        Application.StatusBar = "Step 5: Processing Operation Mode Summary section..."
        
        Dim statusColSummary As Long
        Dim opCodeColSummary As Long
        Dim summaryCount As Long
        
        ' Find columns in summary section
        statusColSummary = FindColumnByHeader(wsEval, summaryRow + 1, "Final Status")
        opCodeColSummary = FindColumnByHeader(wsEval, summaryRow + 1, "Op Code")
        
        If statusColSummary > 0 And opCodeColSummary > 0 Then
            debugMsg = debugMsg & "Processing Operation Mode Summary:" & vbCrLf & _
                       "  - Op Code column: " & ColumnLetter(opCodeColSummary) & " (" & opCodeColSummary & ")" & vbCrLf & _
                       "  - Final Status column: " & ColumnLetter(statusColSummary) & " (" & statusColSummary & ")" & vbCrLf
            
            summaryCount = 0
            
            ' Process each operation in summary section
            For i = summaryRow + 2 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, opCodeColSummary).Value))
                
                ' Stop if empty or next section
                If opCode = "" Then
                    Exit For
                End If
                
                If IsNumeric(opCode) Then
                    summaryCount = summaryCount + 1
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColSummary).Value)))
                    
                    ' Update HeatMap if valid status
                    If finalStatus <> "" And finalStatus <> "N/A" And statusColHeatMap > 0 Then
                        If UpdateOperationInHeatMap(wsHeatMap, opCode, finalStatus, statusColHeatMap, lastRowHeatMap) Then
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
            Next i
            
            debugMsg = debugMsg & "  - Parent operations found: " & summaryCount & vbCrLf & vbCrLf
        Else
            debugMsg = debugMsg & "  - ERROR: Could not find required columns in summary section!" & vbCrLf & vbCrLf
        End If
    End If
    
    ' === STEP 6: Show Results ===
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    debugMsg = debugMsg & "RESULTS:" & vbCrLf & _
               "  - Operations found in Evaluation: " & evalRowsFound & vbCrLf & _
               "  - Operations in HeatMap: " & heatMapRowsFound & vbCrLf & _
               "  - Operations updated: " & updatedCount & vbCrLf & _
               "  - Time elapsed: " & Format(elapsedTime, "0.00") & " seconds"
    
    ' Show comprehensive debug message
    MsgBox debugMsg, vbInformation, "HeatMap Update Complete - Debug Report"
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "At line: " & Erl, _
           vbCritical, "Update Failed"
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

' Helper function to find a column by header name
Private Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim col As Long
    Dim cellValue As String
    
    FindColumnByHeader = 0
    
    For col = 1 To 50 ' Search up to column AX
        cellValue = Trim(UCase(CStr(ws.Cells(headerRow, col).Value)))
        If InStr(1, cellValue, UCase(headerName), vbTextCompare) > 0 Then
            FindColumnByHeader = col
            Exit Function
        End If
    Next col
End Function

' Helper function to update operation status in HeatMap
Private Function UpdateOperationInHeatMap(ws As Worksheet, opCode As String, _
                                          status As String, statusCol As Long, _
                                          lastRow As Long) As Boolean
    Dim i As Long
    Dim heatMapOpCode As String
    Dim statusDot As String
    Dim statusColor As Long
    
    UpdateOperationInHeatMap = False
    
    ' Find the operation in HeatMap
    For i = 2 To lastRow
        heatMapOpCode = Trim(CStr(ws.Cells(i, 1).Value))
        
        If heatMapOpCode = opCode Then
            ' Determine status dot and color
            Select Case UCase(status)
                Case "RED"
                    statusDot = "●"
                    statusColor = RGB(255, 0, 0) ' Red
                Case "YELLOW"
                    statusDot = "●"
                    statusColor = RGB(255, 255, 0) ' Yellow
                Case "GREEN"
                    statusDot = "●"
                    statusColor = RGB(0, 176, 80) ' Green
                Case Else
                    statusDot = "●"
                    statusColor = RGB(128, 128, 128) ' Gray for N/A
            End Select
            
            ' Update the cell
            With ws.Cells(i, statusCol)
                .Value = statusDot
                .Font.Name = "Wingdings"
                .Font.Size = 14
                .Font.Color = statusColor
                .HorizontalAlignment = xlCenter
            End With
            
            UpdateOperationInHeatMap = True
            Exit Function
        End If
    Next i
End Function

' Helper function to convert column number to letter
Private Function ColumnLetter(col As Long) As String
    Dim result As String
    Dim num As Long
    
    num = col
    Do While num > 0
        result = Chr((num - 1) Mod 26 + 65) & result
        num = (num - 1) \ 26
    Loop
    
    ColumnLetter = result
End Function

' Helper function to list all sheet names
Private Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        If sheetList <> "" Then sheetList = sheetList & ", "
        sheetList = sheetList & "'" & ws.Name & "'"
    Next ws
    
    ListAllSheets = sheetList
End Function

' Function to create the Update button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnExists As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets("HeatMap_Sheet")
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Sheets("Heatmap Sheet")
        End If
    End If
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Cannot find HeatMap Sheet to create button!" & vbCrLf & _
               "Please ensure the sheet exists.", vbExclamation
        Exit Sub
    End If
    
    ' Check if button already exists
    btnExists = False
    On Error Resume Next
    For Each btn In ws.Buttons
        If btn.Caption = "Update HeatMap Status" Then
            btnExists = True
            MsgBox "Button already exists on " & ws.Name & "!", vbInformation
            Exit Sub
        End If
    Next btn
    On Error GoTo 0
    
    ' Create the button
    Set btn = ws.Buttons.Add(100, 10, 200, 30)
    btn.Caption = "Update HeatMap Status"
    btn.OnAction = "UpdateHeatMapStatus"
    
    ' Format the button
    With btn
        .Font.Bold = True
        .Font.Size = 11
    End With
    
    MsgBox "Button created successfully on " & ws.Name & "!" & vbCrLf & vbCrLf & _
           "Click the button to transfer evaluation results.", vbInformation
End Sub
