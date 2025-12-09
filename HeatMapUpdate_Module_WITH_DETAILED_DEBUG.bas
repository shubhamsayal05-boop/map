Attribute VB_Name = "HeatMapUpdate_Debug"
' ====================================================================
' Module: HeatMapUpdate_Debug
' Purpose: Transfer evaluation results to HeatMap Sheet with detailed debugging
' ====================================================================

Option Explicit

' Main function to update HeatMap status with comprehensive debugging
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
    Dim evalFound As Long, heatMapFound As Long
    Dim matchedCount As Long
    Dim sampleData As String
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalFound = 0
    heatMapFound = 0
    matchedCount = 0
    debugMsg = "=== HEATMAP STATUS UPDATE DIAGNOSTIC ===" & vbCrLf & vbCrLf
    
    ' Step 1: Find worksheets
    debugMsg = debugMsg & "STEP 1: Finding worksheets..." & vbCrLf
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    If wsEval Is Nothing Then
        MsgBox "ERROR: 'Evaluation Results' sheet not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetSheetNames(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found 'Evaluation Results' sheet" & vbCrLf
    
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    If wsHeatMap Is Nothing Then
        MsgBox "ERROR: 'HeatMap Sheet' not found!" & vbCrLf & vbCrLf & _
               "Available sheets:" & vbCrLf & GetSheetNames(), _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    debugMsg = debugMsg & "✓ Found 'HeatMap Sheet'" & vbCrLf & vbCrLf
    On Error GoTo ErrorHandler
    
    Application.ScreenUpdating = False
    Application.StatusBar = "Analyzing sheets..."
    
    ' Step 2: Find last rows
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = debugMsg & "STEP 2: Sheet dimensions:" & vbCrLf
    debugMsg = debugMsg & "- Evaluation Results: " & lastRowEval & " rows" & vbCrLf
    debugMsg = debugMsg & "- HeatMap Sheet: " & lastRowHeatMap & " rows" & vbCrLf & vbCrLf
    
    ' Step 3: Find "Overall Status by Op Code" section
    debugMsg = debugMsg & "STEP 3: Looking for 'Overall Status by Op Code'..." & vbCrLf
    Dim overallStartRow As Long
    overallStartRow = 0
    
    For i = 1 To lastRowEval
        If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallStartRow = i
            Exit For
        End If
    Next i
    
    If overallStartRow = 0 Then
        debugMsg = debugMsg & "✗ 'Overall Status by Op Code' section NOT FOUND!" & vbCrLf
        debugMsg = debugMsg & "Showing first 10 rows of column A:" & vbCrLf
        For i = 1 To Application.Min(10, lastRowEval)
            debugMsg = debugMsg & "  Row " & i & ": " & CStr(wsEval.Cells(i, 1).Value) & vbCrLf
        Next i
        MsgBox debugMsg, vbExclamation, "Section Not Found"
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If
    
    debugMsg = debugMsg & "✓ Found at row " & overallStartRow & vbCrLf & vbCrLf
    
    ' Step 4: Find "Final Status" column
    debugMsg = debugMsg & "STEP 4: Finding 'Final Status' column..." & vbCrLf
    Dim statusColEval As Long
    Dim headerRow As Long
    headerRow = overallStartRow + 1
    
    statusColEval = 0
    For j = 1 To 20 ' Check first 20 columns
        If InStr(1, CStr(wsEval.Cells(headerRow, j).Value), "Final Status", vbTextCompare) > 0 Then
            statusColEval = j
            Exit For
        End If
    Next j
    
    If statusColEval = 0 Then
        debugMsg = debugMsg & "✗ 'Final Status' column NOT FOUND!" & vbCrLf
        debugMsg = debugMsg & "Header row " & headerRow & " columns:" & vbCrLf
        For j = 1 To Application.Min(15, wsEval.Cells(headerRow, wsEval.Columns.Count).End(xlToLeft).Column)
            debugMsg = debugMsg & "  Col " & j & ": " & CStr(wsEval.Cells(headerRow, j).Value) & vbCrLf
        Next j
        MsgBox debugMsg, vbExclamation, "Column Not Found"
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If
    
    debugMsg = debugMsg & "✓ Found at column " & statusColEval & " (" & ColumnLetter(statusColEval) & ")" & vbCrLf & vbCrLf
    
    ' Step 5: Read evaluation data
    debugMsg = debugMsg & "STEP 5: Reading evaluation data..." & vbCrLf
    Application.StatusBar = "Reading evaluation data..."
    
    Dim dataStartRow As Long
    dataStartRow = headerRow + 1
    
    sampleData = ""
    For i = dataStartRow To Application.Min(dataStartRow + 4, lastRowEval)
        opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
        
        ' Stop if we hit the next section
        If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
            Exit For
        End If
        
        If opCode <> "" And IsNumeric(opCode) Then
            evalFound = evalFound + 1
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColEval).Value)))
            
            If sampleData = "" Then
                sampleData = "Sample evaluation data:" & vbCrLf
            End If
            If evalFound <= 5 Then
                sampleData = sampleData & "  OpCode: " & opCode & " → Status: " & finalStatus & vbCrLf
            End If
        End If
    Next i
    
    debugMsg = debugMsg & "✓ Found " & evalFound & " operation codes with status" & vbCrLf
    debugMsg = debugMsg & sampleData & vbCrLf
    
    ' Step 6: Check HeatMap structure
    debugMsg = debugMsg & "STEP 6: Checking HeatMap Sheet structure..." & vbCrLf
    
    ' Find "Status" column in HeatMap
    Dim statusColHeatMap As Long
    statusColHeatMap = 0
    
    ' Check first row for "Status" header
    For j = 1 To 20
        If InStr(1, CStr(wsHeatMap.Cells(1, j).Value), "Status", vbTextCompare) > 0 Then
            statusColHeatMap = j
            Exit For
        End If
    Next j
    
    If statusColHeatMap = 0 Then
        debugMsg = debugMsg & "✗ 'Status' column NOT FOUND in HeatMap!" & vbCrLf
        debugMsg = debugMsg & "Row 1 columns:" & vbCrLf
        For j = 1 To Application.Min(10, wsHeatMap.Cells(1, wsHeatMap.Columns.Count).End(xlToLeft).Column)
            debugMsg = debugMsg & "  Col " & j & ": " & CStr(wsHeatMap.Cells(1, j).Value) & vbCrLf
        Next j
        MsgBox debugMsg, vbExclamation, "Status Column Not Found"
        Application.ScreenUpdating = True
        Application.StatusBar = False
        Exit Sub
    End If
    
    debugMsg = debugMsg & "✓ 'Status' column found at column " & statusColHeatMap & " (" & ColumnLetter(statusColHeatMap) & ")" & vbCrLf
    
    ' Count HeatMap operations
    For i = 2 To lastRowHeatMap
        opCode = Trim(CStr(wsHeatMap.Cells(i, 1).Value))
        If opCode <> "" And IsNumeric(opCode) Then
            heatMapFound = heatMapFound + 1
        End If
    Next i
    
    debugMsg = debugMsg & "✓ Found " & heatMapFound & " operation codes in HeatMap" & vbCrLf & vbCrLf
    
    ' Step 7: Update statuses
    debugMsg = debugMsg & "STEP 7: Updating statuses..." & vbCrLf
    Application.StatusBar = "Updating HeatMap statuses..."
    
    Dim firstMatch As String
    firstMatch = ""
    
    For i = dataStartRow To lastRowEval
        opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
        
        ' Stop if we hit the next section
        If InStr(1, opCode, "Operation Mode Summary", vbTextCompare) > 0 Then
            Exit For
        End If
        
        If opCode <> "" And IsNumeric(opCode) Then
            finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusColEval).Value)))
            
            ' Skip if no valid status
            If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                ' Find in HeatMap
                For j = 2 To lastRowHeatMap
                    If Trim(CStr(wsHeatMap.Cells(j, 1).Value)) = opCode Then
                        matchedCount = matchedCount + 1
                        
                        ' Update status with colored dot
                        Dim dotChar As String
                        Dim dotColor As Long
                        dotChar = "●"
                        
                        Select Case finalStatus
                            Case "RED"
                                dotColor = RGB(255, 0, 0)
                            Case "YELLOW"
                                dotColor = RGB(255, 255, 0)
                            Case "GREEN"
                                dotColor = RGB(0, 176, 80)
                            Case Else
                                dotColor = RGB(128, 128, 128) ' Gray
                        End Select
                        
                        With wsHeatMap.Cells(j, statusColHeatMap)
                            .Value = dotChar
                            .Font.Name = "Wingdings"
                            .Font.Size = 14
                            .Font.Color = dotColor
                        End With
                        
                        updatedCount = updatedCount + 1
                        
                        If firstMatch = "" Then
                            firstMatch = "First match: OpCode " & opCode & " → " & finalStatus & " (Row " & j & ")" & vbCrLf
                        End If
                        
                        Exit For
                    End If
                Next j
            End If
        End If
    Next i
    
    debugMsg = debugMsg & "✓ Matched " & matchedCount & " operations" & vbCrLf
    debugMsg = debugMsg & "✓ Updated " & updatedCount & " statuses" & vbCrLf
    If firstMatch <> "" Then
        debugMsg = debugMsg & firstMatch
    End If
    debugMsg = debugMsg & vbCrLf
    
    ' Step 8: Process "Operation Mode Summary" section
    debugMsg = debugMsg & "STEP 8: Processing 'Operation Mode Summary'..." & vbCrLf
    
    Dim summaryStartRow As Long
    summaryStartRow = 0
    
    For i = overallStartRow To lastRowEval
        If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
            summaryStartRow = i
            Exit For
        End If
    Next i
    
    If summaryStartRow > 0 Then
        debugMsg = debugMsg & "✓ Found at row " & summaryStartRow & vbCrLf
        
        ' Find Final Status column in summary header
        Dim summaryHeaderRow As Long
        Dim summaryStatusCol As Long
        Dim summaryOpCodeCol As Long
        
        summaryHeaderRow = summaryStartRow + 1
        summaryStatusCol = 0
        summaryOpCodeCol = 0
        
        For j = 1 To 20
            If InStr(1, CStr(wsEval.Cells(summaryHeaderRow, j).Value), "Final Status", vbTextCompare) > 0 Then
                summaryStatusCol = j
            End If
            If InStr(1, CStr(wsEval.Cells(summaryHeaderRow, j).Value), "Op Code", vbTextCompare) > 0 Then
                summaryOpCodeCol = j
            End If
        Next j
        
        If summaryOpCodeCol > 0 And summaryStatusCol > 0 Then
            Dim summaryUpdated As Long
            summaryUpdated = 0
            
            For i = summaryHeaderRow + 1 To lastRowEval
                opCode = Trim(CStr(wsEval.Cells(i, summaryOpCodeCol).Value))
                
                If opCode <> "" And IsNumeric(opCode) Then
                    finalStatus = Trim(UCase(CStr(wsEval.Cells(i, summaryStatusCol).Value)))
                    
                    If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                        ' Find in HeatMap
                        For j = 2 To lastRowHeatMap
                            If Trim(CStr(wsHeatMap.Cells(j, 1).Value)) = opCode Then
                                ' Update if not already updated
                                Dim currentVal As String
                                currentVal = CStr(wsHeatMap.Cells(j, statusColHeatMap).Value)
                                
                                If currentVal = "" Or currentVal = "●" Then
                                    Dim dotCharSum As String
                                    Dim dotColorSum As Long
                                    dotCharSum = "●"
                                    
                                    Select Case finalStatus
                                        Case "RED"
                                            dotColorSum = RGB(255, 0, 0)
                                        Case "YELLOW"
                                            dotColorSum = RGB(255, 255, 0)
                                        Case "GREEN"
                                            dotColorSum = RGB(0, 176, 80)
                                        Case Else
                                            dotColorSum = RGB(128, 128, 128)
                                    End Select
                                    
                                    With wsHeatMap.Cells(j, statusColHeatMap)
                                        .Value = dotCharSum
                                        .Font.Name = "Wingdings"
                                        .Font.Size = 14
                                        .Font.Color = dotColorSum
                                    End With
                                    
                                    summaryUpdated = summaryUpdated + 1
                                End If
                                Exit For
                            End If
                        Next j
                    End If
                End If
            Next i
            
            debugMsg = debugMsg & "✓ Updated " & summaryUpdated & " parent operation statuses" & vbCrLf
            updatedCount = updatedCount + summaryUpdated
        Else
            debugMsg = debugMsg & "✗ Could not find columns in summary section" & vbCrLf
        End If
    Else
        debugMsg = debugMsg & "✗ 'Operation Mode Summary' section NOT found" & vbCrLf
    End If
    
    ' Final summary
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    
    debugMsg = debugMsg & vbCrLf & "=== SUMMARY ===" & vbCrLf
    debugMsg = debugMsg & "✓ Total operations updated: " & updatedCount & vbCrLf
    debugMsg = debugMsg & "✓ Time taken: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf
    
    If updatedCount = 0 Then
        debugMsg = debugMsg & vbCrLf & "⚠ NO STATUSES WERE UPDATED!" & vbCrLf & _
                   "Possible reasons:" & vbCrLf & _
                   "1. Op Codes don't match between sheets" & vbCrLf & _
                   "2. No valid statuses (all N/A)" & vbCrLf & _
                   "3. Status column location incorrect"
        MsgBox debugMsg, vbExclamation, "HeatMap Update - No Changes"
    Else
        MsgBox debugMsg, vbInformation, "HeatMap Update Complete"
    End If
    
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "ERROR: " & Err.Description & vbCrLf & vbCrLf & _
           "At line: " & Erl & vbCrLf & vbCrLf & _
           "Debug info:" & vbCrLf & debugMsg, _
           vbCritical, "Update Failed"
End Sub

' Helper function to get sheet names
Function GetSheetNames() As String
    Dim ws As Worksheet
    Dim names As String
    names = ""
    For Each ws In ThisWorkbook.Worksheets
        names = names & "- " & ws.Name & vbCrLf
    Next ws
    GetSheetNames = names
End Function

' Helper function to convert column number to letter
Function ColumnLetter(colNum As Long) As String
    Dim temp As Long
    Dim letter As String
    letter = ""
    temp = colNum
    Do While temp > 0
        temp = temp - 1
        letter = Chr(65 + (temp Mod 26)) & letter
        temp = temp \ 26
    Loop
    ColumnLetter = letter
End Function

' Create Update Button on HeatMap Sheet
Sub CreateUpdateButton()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnLeft As Double
    Dim btnTop As Double
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found!", vbExclamation
        Exit Sub
    End If
    
    ' Delete existing button if present
    On Error Resume Next
    ws.Buttons("UpdateHeatMapBtn").Delete
    On Error GoTo 0
    
    ' Create button in top-right area
    btnLeft = ws.Range("O2").Left
    btnTop = ws.Range("O2").Top
    
    Set btn = ws.Buttons.Add(btnLeft, btnTop, 180, 30)
    With btn
        .Name = "UpdateHeatMapBtn"
        .Caption = "Update HeatMap Status"
        .OnAction = "UpdateHeatMapStatus"
        .Font.Bold = True
        .Font.Size = 10
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button after running evaluation to update statuses.", _
           vbInformation, "Button Created"
End Sub
