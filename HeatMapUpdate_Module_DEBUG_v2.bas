Attribute VB_Name = "HeatMapUpdate_DEBUG"
' ====================================================================
' Module: HeatMapUpdate_DEBUG (Enhanced Version)
' Purpose: Transfer evaluation results to HeatMap Sheet with detailed debugging
' New Requirement: Works with "Overall Status by Op Code" and 
'                  "Operation Mode Summary" sections in Evaluation Results
' ====================================================================

Option Explicit

' Main function to update HeatMap status from evaluation results with debugging
Sub UpdateHeatMapStatus_DEBUG()
    Dim wsEval As Worksheet
    Dim wsHeatMap As Worksheet
    Dim lastRowEval As Long
    Dim lastRowHeatMap As Long
    Dim i As Long
    Dim opCode As String
    Dim finalStatus As String
    Dim updatedCount As Long
    Dim startTime As Double
    Dim debugMsg As String
    Dim overallStartRow As Long
    Dim summaryStartRow As Long
    Dim foundSheets As Boolean
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    foundSheets = True
    debugMsg = "=== HeatMap Update Debug Report ===" & vbCrLf & vbCrLf
    
    ' Step 1: Check if sheets exist
    debugMsg = debugMsg & "STEP 1: Checking for required sheets..." & vbCrLf
    
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        debugMsg = debugMsg & "❌ ERROR: 'Evaluation Results' sheet not found!" & vbCrLf
        debugMsg = debugMsg & "   Available sheets: " & ListAllSheets() & vbCrLf
        foundSheets = False
    Else
        debugMsg = debugMsg & "✓ Found 'Evaluation Results' sheet" & vbCrLf
    End If
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        debugMsg = debugMsg & "❌ ERROR: 'HeatMap Sheet' not found!" & vbCrLf
        debugMsg = debugMsg & "   Available sheets: " & ListAllSheets() & vbCrLf
        foundSheets = False
    Else
        debugMsg = debugMsg & "✓ Found 'HeatMap Sheet'" & vbCrLf
    End If
    On Error GoTo ErrorHandler
    
    If Not foundSheets Then
        MsgBox debugMsg, vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' Step 2: Find data sections in Evaluation Results
    debugMsg = debugMsg & "STEP 2: Locating data sections in Evaluation Results..." & vbCrLf
    
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "   Total rows in Evaluation Results: " & lastRowEval & vbCrLf
    
    ' Find "Overall Status by Op Code" section
    overallStartRow = FindSectionRow(wsEval, "Overall Status by Op Code")
    If overallStartRow > 0 Then
        debugMsg = debugMsg & "✓ Found 'Overall Status by Op Code' at row " & overallStartRow & vbCrLf
    Else
        debugMsg = debugMsg & "❌ WARNING: 'Overall Status by Op Code' section not found" & vbCrLf
    End If
    
    ' Find "Operation Mode Summary" section
    summaryStartRow = FindSectionRow(wsEval, "Operation Mode Summary")
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "✓ Found 'Operation Mode Summary' at row " & summaryStartRow & vbCrLf
    Else
        debugMsg = debugMsg & "❌ WARNING: 'Operation Mode Summary' section not found" & vbCrLf
    End If
    
    debugMsg = debugMsg & vbCrLf
    
    ' Step 3: Check HeatMap Sheet structure
    debugMsg = debugMsg & "STEP 3: Checking HeatMap Sheet structure..." & vbCrLf
    
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    debugMsg = debugMsg & "   Total rows in HeatMap Sheet: " & lastRowHeatMap & vbCrLf
    
    ' Check column A header
    Dim colAHeader As String
    colAHeader = Trim(CStr(wsHeatMap.Cells(1, 1).Value))
    debugMsg = debugMsg & "   Column A header: '" & colAHeader & "'" & vbCrLf
    
    ' Find Status column
    Dim statusCol As Long
    statusCol = FindStatusColumn(wsHeatMap)
    If statusCol > 0 Then
        Dim statusHeader As String
        statusHeader = Trim(CStr(wsHeatMap.Cells(1, statusCol).Value))
        debugMsg = debugMsg & "✓ Found Status column at column " & statusCol & " ('" & statusHeader & "')" & vbCrLf
    Else
        debugMsg = debugMsg & "❌ ERROR: Could not find Status column in HeatMap Sheet!" & vbCrLf
        debugMsg = debugMsg & "   Row 1 headers: " & ListRowHeaders(wsHeatMap, 1) & vbCrLf
        MsgBox debugMsg, vbCritical, "Missing Status Column"
        Exit Sub
    End If
    
    ' Sample first few operation codes from HeatMap
    debugMsg = debugMsg & "   Sample Op Codes from HeatMap (rows 2-5):" & vbCrLf
    For i = 2 To Application.Min(5, lastRowHeatMap)
        Dim sampleCode As String
        sampleCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If sampleCode <> "" Then
            debugMsg = debugMsg & "      Row " & i & ": " & sampleCode & vbCrLf
        End If
    Next i
    
    debugMsg = debugMsg & vbCrLf
    
    ' Step 4: Process "Overall Status by Op Code" section
    If overallStartRow > 0 Then
        debugMsg = debugMsg & "STEP 4a: Processing 'Overall Status by Op Code' section..." & vbCrLf
        
        Dim overallCount As Long
        overallCount = 0
        
        ' Find the header row (next row after section title)
        Dim headerRow As Long
        headerRow = overallStartRow + 1
        
        ' Find Op Code and Overall Status columns
        Dim opCodeCol As Long, overallStatusCol As Long
        opCodeCol = FindColumnByHeader(wsEval, headerRow, "Op Code")
        overallStatusCol = FindColumnByHeader(wsEval, headerRow, "Overall Status")
        
        If opCodeCol > 0 And overallStatusCol > 0 Then
            debugMsg = debugMsg & "   Op Code column: " & opCodeCol & vbCrLf
            debugMsg = debugMsg & "   Overall Status column: " & overallStatusCol & vbCrLf
            
            ' Process rows until we hit another section or empty rows
            i = headerRow + 1
            Do While i <= lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, opCodeCol).Value))
                
                ' Stop if we hit a new section or empty code
                If opCode = "" Or InStr(1, opCode, "Summary", vbTextCompare) > 0 Or _
                   InStr(1, opCode, "Accelerations", vbTextCompare) > 0 Then
                    Exit Do
                End If
                
                If IsNumeric(opCode) Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, overallStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        If UpdateOperationStatus_DEBUG(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap, debugMsg) Then
                            overallCount = overallCount + 1
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
                
                i = i + 1
            Loop
            
            debugMsg = debugMsg & "   Processed " & overallCount & " operations from Overall Status section" & vbCrLf
        Else
            debugMsg = debugMsg & "❌ Could not find required columns in Overall Status section" & vbCrLf
        End If
        
        debugMsg = debugMsg & vbCrLf
    End If
    
    ' Step 5: Process "Operation Mode Summary" section
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "STEP 4b: Processing 'Operation Mode Summary' section..." & vbCrLf
        
        Dim summaryCount As Long
        summaryCount = 0
        
        ' Find the header row
        headerRow = summaryStartRow + 1
        
        ' Find Op Code and Final Status columns
        opCodeCol = FindColumnByHeader(wsEval, headerRow, "Op Code")
        Dim finalStatusCol As Long
        finalStatusCol = FindColumnByHeader(wsEval, headerRow, "Final Status")
        
        If opCodeCol > 0 And finalStatusCol > 0 Then
            debugMsg = debugMsg & "   Op Code column: " & opCodeCol & vbCrLf
            debugMsg = debugMsg & "   Final Status column: " & finalStatusCol & vbCrLf
            
            ' Process rows
            i = headerRow + 1
            Do While i <= lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, opCodeCol).Value))
                
                ' Stop if empty or new section
                If opCode = "" Then Exit Do
                
                If IsNumeric(opCode) Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, finalStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "N/A" Then
                        If UpdateOperationStatus_DEBUG(wsHeatMap, opCode, finalStatus, statusCol, lastRowHeatMap, debugMsg) Then
                            summaryCount = summaryCount + 1
                            updatedCount = updatedCount + 1
                        End If
                    End If
                End If
                
                i = i + 1
            Loop
            
            debugMsg = debugMsg & "   Processed " & summaryCount & " operations from Operation Mode Summary" & vbCrLf
        Else
            debugMsg = debugMsg & "❌ Could not find required columns in Operation Mode Summary section" & vbCrLf
        End If
        
        debugMsg = debugMsg & vbCrLf
    End If
    
    ' Final summary
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    debugMsg = debugMsg & "=== FINAL RESULTS ===" & vbCrLf
    debugMsg = debugMsg & "Total operations updated: " & updatedCount & vbCrLf
    debugMsg = debugMsg & "Time taken: " & Format(Timer - startTime, "0.00") & " seconds" & vbCrLf
    
    If updatedCount = 0 Then
        debugMsg = debugMsg & vbCrLf & "❌ NO OPERATIONS WERE UPDATED!" & vbCrLf
        debugMsg = debugMsg & "Please check the structure of your sheets and ensure:" & vbCrLf
        debugMsg = debugMsg & "1. Evaluation Results has 'Overall Status by Op Code' or 'Operation Mode Summary' sections" & vbCrLf
        debugMsg = debugMsg & "2. HeatMap Sheet has a 'Status' column" & vbCrLf
        debugMsg = debugMsg & "3. Op Codes match between the two sheets" & vbCrLf
        MsgBox debugMsg, vbExclamation, "HeatMap Update - Debug Report"
    Else
        debugMsg = debugMsg & vbCrLf & "✓ Update completed successfully!" & vbCrLf
        MsgBox debugMsg, vbInformation, "HeatMap Update - Success"
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    debugMsg = debugMsg & vbCrLf & "❌ ERROR: " & Err.Description & vbCrLf
    debugMsg = debugMsg & "Error Number: " & Err.Number & vbCrLf
    MsgBox debugMsg, vbCritical, "HeatMap Update Error"
End Sub

' Helper: Find a section row by searching for section title in column A or nearby
Function FindSectionRow(ws As Worksheet, sectionTitle As String) As Long
    Dim i As Long
    Dim cellValue As String
    Dim lastRow As Long
    
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For i = 1 To lastRow
        cellValue = Trim(CStr(ws.Cells(i, 1).Value))
        If InStr(1, cellValue, sectionTitle, vbTextCompare) > 0 Then
            FindSectionRow = i
            Exit Function
        End If
        
        ' Also check a few columns to the right
        If cellValue = "" Then
            cellValue = Trim(CStr(ws.Cells(i, 2).Value))
            If InStr(1, cellValue, sectionTitle, vbTextCompare) > 0 Then
                FindSectionRow = i
                Exit Function
            End If
        End If
    Next i
    
    FindSectionRow = 0
End Function

' Helper: Find column by header name in a specific row
Function FindColumnByHeader(ws As Worksheet, headerRow As Long, headerName As String) As Long
    Dim col As Long
    Dim cellValue As String
    
    For col = 1 To 50 ' Check first 50 columns
        cellValue = Trim(CStr(ws.Cells(headerRow, col).Value))
        If InStr(1, cellValue, headerName, vbTextCompare) > 0 Then
            FindColumnByHeader = col
            Exit Function
        End If
    Next col
    
    FindColumnByHeader = 0
End Function

' Helper: Find Status column in HeatMap Sheet
Function FindStatusColumn(ws As Worksheet) As Long
    Dim col As Long
    Dim cellValue As String
    
    For col = 1 To 50
        cellValue = Trim(CStr(ws.Cells(1, col).Value))
        If InStr(1, cellValue, "Status", vbTextCompare) > 0 And _
           InStr(1, cellValue, "Current", vbTextCompare) > 0 Then
            FindStatusColumn = col
            Exit Function
        ElseIf cellValue = "Status" Or cellValue = "status" Then
            FindStatusColumn = col
            Exit Function
        End If
    Next col
    
    FindStatusColumn = 0
End Function

' Helper: Update operation status in HeatMap with debugging
Function UpdateOperationStatus_DEBUG(wsHeatMap As Worksheet, opCode As String, _
                                     status As String, statusCol As Long, _
                                     lastRow As Long, ByRef debugMsg As String) As Boolean
    Dim i As Long
    Dim heatMapCode As String
    
    For i = 2 To lastRow
        heatMapCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        
        If heatMapCode = opCode Then
            ' Set status dot with color
            Call SetStatusDot(wsHeatMap.Cells(i, statusCol), status)
            UpdateOperationStatus_DEBUG = True
            Exit Function
        End If
    Next i
    
    ' Not found - add to debug log
    debugMsg = debugMsg & "   ⚠ Op Code " & opCode & " not found in HeatMap Sheet" & vbCrLf
    UpdateOperationStatus_DEBUG = False
End Function

' Helper: Set status dot with appropriate color
Sub SetStatusDot(cell As Range, status As String)
    Dim dotChar As String
    Dim dotColor As Long
    
    ' Use filled circle character
    dotChar = ChrW(&H25CF) ' Filled circle: ●
    
    ' Set color based on status
    Select Case status
        Case "RED"
            dotColor = RGB(255, 0, 0)     ' Red
        Case "YELLOW"
            dotColor = RGB(255, 192, 0)   ' Yellow/Orange
        Case "GREEN"
            dotColor = RGB(0, 176, 80)    ' Green
        Case Else
            dotColor = RGB(128, 128, 128) ' Gray for N/A
            dotChar = "N/A"
    End Select
    
    ' Set the cell value and color
    With cell
        .Value = dotChar
        .Font.Color = dotColor
        .Font.Name = "Wingdings"
        .Font.Size = 14
        .HorizontalAlignment = xlCenter
    End With
End Sub

' Helper: List all sheet names
Function ListAllSheets() As String
    Dim ws As Worksheet
    Dim sheetList As String
    
    sheetList = ""
    For Each ws In ThisWorkbook.Worksheets
        If sheetList <> "" Then sheetList = sheetList & ", "
        sheetList = sheetList & """" & ws.Name & """"
    Next ws
    
    ListAllSheets = sheetList
End Function

' Helper: List headers in a row
Function ListRowHeaders(ws As Worksheet, rowNum As Long) As String
    Dim col As Long
    Dim headerList As String
    Dim cellValue As String
    
    headerList = ""
    For col = 1 To 20
        cellValue = Trim(CStr(ws.Cells(rowNum, col).Value))
        If cellValue <> "" Then
            If headerList <> "" Then headerList = headerList & ", "
            headerList = headerList & "Col" & col & ": """ & cellValue & """"
        End If
    Next col
    
    ListRowHeaders = headerList
End Function

' Create button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnName As String
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found! Please create it first.", vbCritical
        Exit Sub
    End If
    
    ' Delete existing button if any
    btnName = "btnUpdateHeatMap_DEBUG"
    ws.Buttons(btnName).Delete
    
    ' Create new button
    Set btn = ws.Buttons.Add(10, 10, 200, 30)
    With btn
        .Name = btnName
        .Caption = "Update HeatMap Status (DEBUG)"
        .OnAction = "UpdateHeatMapStatus_DEBUG"
    End With
    
    On Error GoTo 0
    
    MsgBox "Debug button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update statuses with detailed debug information.", _
           vbInformation, "Button Created"
End Sub
