Attribute VB_Name = "HeatMapUpdateEnhancedDebug"
' ====================================================================
' Module: HeatMapUpdateEnhancedDebug
' Purpose: Transfer evaluation results to HeatMap Sheet with comprehensive debugging
' Version: Enhanced Final Debug - Detailed error messages
' ====================================================================

Option Explicit

' Main function with detailed debugging and error messages
Sub UpdateHeatMapStatusWithDebug()
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
    Dim evalSheetFound As Boolean, heatMapSheetFound As Boolean
    Dim overallSectionFound As Boolean, summarySectionFound As Boolean
    Dim statusCol As Long
    Dim heatMapStatusCol As Long
    
    On Error GoTo ErrorHandler
    
    startTime = Timer
    updatedCount = 0
    evalSheetFound = False
    heatMapSheetFound = False
    overallSectionFound = False
    summarySectionFound = False
    debugMsg = "=== HeatMap Update Debug Report ===" & vbCrLf & vbCrLf
    
    ' Step 1: Check for Evaluation Results sheet
    Application.StatusBar = "Step 1/7: Checking Evaluation Results sheet..."
    On Error Resume Next
    Set wsEval = ThisWorkbook.Sheets("Evaluation Results")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        debugMsg = debugMsg & "❌ FAILED: 'Evaluation Results' sheet not found" & vbCrLf & vbCrLf
        debugMsg = debugMsg & "Available sheets in workbook:" & vbCrLf
        Dim ws As Worksheet
        For Each ws In ThisWorkbook.Sheets
            debugMsg = debugMsg & "  - " & ws.Name & vbCrLf
        Next ws
        debugMsg = debugMsg & vbCrLf & "ACTION: Rename your evaluation sheet to exactly 'Evaluation Results'"
        MsgBox debugMsg, vbCritical, "Sheet Not Found"
        Exit Sub
    Else
        evalSheetFound = True
        debugMsg = debugMsg & "✓ Step 1: Found 'Evaluation Results' sheet" & vbCrLf
    End If
    
    ' Step 2: Check for HeatMap Sheet
    Application.StatusBar = "Step 2/7: Checking HeatMap Sheet..."
    On Error Resume Next
    Set wsHeatMap = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsHeatMap Is Nothing Then
        debugMsg = debugMsg & "❌ FAILED: 'HeatMap Sheet' not found" & vbCrLf & vbCrLf
        debugMsg = debugMsg & "Available sheets in workbook:" & vbCrLf
        For Each ws In ThisWorkbook.Sheets
            debugMsg = debugMsg & "  - " & ws.Name & vbCrLf
        Next ws
        debugMsg = debugMsg & vbCrLf & "ACTION: Rename your heat map sheet to exactly 'HeatMap Sheet'"
        MsgBox debugMsg, vbCritical, "Sheet Not Found"
        Exit Sub
    Else
        heatMapSheetFound = True
        debugMsg = debugMsg & "✓ Step 2: Found 'HeatMap Sheet'" & vbCrLf
    End If
    
    ' Step 3: Analyze sheet structure
    Application.StatusBar = "Step 3/7: Analyzing sheet structure..."
    lastRowEval = wsEval.Cells(wsEval.Rows.Count, "A").End(xlUp).Row
    lastRowHeatMap = wsHeatMap.Cells(wsHeatMap.Rows.Count, "A").End(xlUp).Row
    
    debugMsg = debugMsg & "✓ Step 3: Sheet structure analyzed" & vbCrLf
    debugMsg = debugMsg & "  - Evaluation Results has " & lastRowEval & " rows" & vbCrLf
    debugMsg = debugMsg & "  - HeatMap Sheet has " & lastRowHeatMap & " rows" & vbCrLf & vbCrLf
    
    ' Step 4: Find "Overall Status by Op Code" section
    Application.StatusBar = "Step 4/7: Searching for 'Overall Status by Op Code' section..."
    Dim overallStartRow As Long
    overallStartRow = 0
    
    For i = 1 To lastRowEval
        Dim cellValue As String
        cellValue = Trim(CStr(wsEval.Cells(i, 1).Value))
        If InStr(1, cellValue, "Overall Status by Op Code", vbTextCompare) > 0 Then
            overallStartRow = i
            overallSectionFound = True
            Exit For
        End If
    Next i
    
    If overallSectionFound Then
        debugMsg = debugMsg & "✓ Step 4: Found 'Overall Status by Op Code' at row " & overallStartRow & vbCrLf
        
        ' Show first few rows of this section
        debugMsg = debugMsg & "  First 3 rows of section:" & vbCrLf
        For i = overallStartRow To Application.Min(overallStartRow + 2, lastRowEval)
            debugMsg = debugMsg & "    Row " & i & ": " & _
                       Left(wsEval.Cells(i, 1).Value & " | " & _
                       wsEval.Cells(i, 2).Value & " | " & _
                       wsEval.Cells(i, 3).Value, 80) & vbCrLf
        Next i
    Else
        debugMsg = debugMsg & "❌ FAILED: 'Overall Status by Op Code' section not found" & vbCrLf
        debugMsg = debugMsg & "  Searched rows 1 to " & lastRowEval & vbCrLf
        debugMsg = debugMsg & "  First 10 values in column A:" & vbCrLf
        For i = 1 To Application.Min(10, lastRowEval)
            debugMsg = debugMsg & "    Row " & i & ": " & wsEval.Cells(i, 1).Value & vbCrLf
        Next i
    End If
    debugMsg = debugMsg & vbCrLf
    
    ' Step 5: Find status column in HeatMap
    Application.StatusBar = "Step 5/7: Finding status column in HeatMap..."
    heatMapStatusCol = 0
    
    ' Look for "Status" or "Current Status" in row 1 of HeatMap
    For j = 1 To 10 ' Check first 10 columns
        cellValue = Trim(UCase(CStr(wsHeatMap.Cells(1, j).Value)))
        If InStr(1, cellValue, "STATUS", vbTextCompare) > 0 Then
            heatMapStatusCol = j
            Exit For
        End If
    Next j
    
    If heatMapStatusCol > 0 Then
        debugMsg = debugMsg & "✓ Step 5: Found Status column in HeatMap at column " & heatMapStatusCol & _
                   " (" & Split(wsHeatMap.Cells(1, heatMapStatusCol).Address, "$")(1) & ")" & vbCrLf
    Else
        debugMsg = debugMsg & "⚠ Step 5: Could not find Status column, using column C (3) as default" & vbCrLf
        heatMapStatusCol = 3 ' Default to column C
    End If
    
    ' Show HeatMap structure
    debugMsg = debugMsg & "  HeatMap header row:" & vbCrLf
    debugMsg = debugMsg & "    Col A: " & wsHeatMap.Cells(1, 1).Value & vbCrLf
    debugMsg = debugMsg & "    Col B: " & wsHeatMap.Cells(1, 2).Value & vbCrLf
    debugMsg = debugMsg & "    Col C: " & wsHeatMap.Cells(1, 3).Value & vbCrLf & vbCrLf
    
    ' Step 6: Process "Overall Status by Op Code" section
    If overallSectionFound Then
        Application.StatusBar = "Step 6/7: Processing operations..."
        
        ' Find Final Status column
        statusCol = 0
        For j = 1 To 20 ' Check first 20 columns in header row
            cellValue = Trim(UCase(CStr(wsEval.Cells(overallStartRow + 1, j).Value)))
            If InStr(1, cellValue, "FINAL STATUS", vbTextCompare) > 0 Or _
               InStr(1, cellValue, "OVERALL STATUS", vbTextCompare) > 0 Then
                statusCol = j
                Exit For
            End If
        Next j
        
        If statusCol = 0 Then
            debugMsg = debugMsg & "⚠ Step 6: Could not find Final Status column, using column C (3)" & vbCrLf
            statusCol = 3 ' Default
        Else
            debugMsg = debugMsg & "✓ Step 6: Found Final Status column at column " & statusCol & vbCrLf
        End If
        
        Dim processedOps As Long, matchedOps As Long
        processedOps = 0
        matchedOps = 0
        
        ' Process each operation
        For i = overallStartRow + 2 To lastRowEval
            opCode = Trim(CStr(wsEval.Cells(i, 1).Value))
            
            ' Stop if we hit next section
            If InStr(1, CStr(wsEval.Cells(i, 1).Value), "Operation Mode Summary", vbTextCompare) > 0 Then
                summarySectionFound = True
                Exit For
            End If
            
            If opCode <> "" And IsNumeric(opCode) And Len(opCode) = 8 Then
                processedOps = processedOps + 1
                finalStatus = Trim(UCase(CStr(wsEval.Cells(i, statusCol).Value)))
                
                If finalStatus <> "" And finalStatus <> "FINAL STATUS" And finalStatus <> "N/A" Then
                    ' Find and update in HeatMap
                    For j = 2 To lastRowHeatMap ' Start from row 2 (skip header)
                        Dim heatMapCode As String
                        heatMapCode = Trim(CStr(wsHeatMap.Cells(j, 1).Value))
                        
                        If heatMapCode = opCode Then
                            matchedOps = matchedOps + 1
                            ' Update status with colored dot
                            Dim dotChar As String
                            Dim dotColor As Long
                            
                            dotChar = "●" ' Filled circle
                            
                            Select Case finalStatus
                                Case "RED"
                                    dotColor = RGB(255, 0, 0)
                                Case "YELLOW"
                                    dotColor = RGB(255, 255, 0)
                                Case "GREEN"
                                    dotColor = RGB(0, 255, 0)
                                Case Else
                                    dotColor = RGB(128, 128, 128) ' Gray for unknown
                            End Select
                            
                            With wsHeatMap.Cells(j, heatMapStatusCol)
                                .Value = dotChar
                                .Font.Name = "Wingdings"
                                .Font.Size = 14
                                .Font.Color = dotColor
                            End With
                            
                            updatedCount = updatedCount + 1
                            Exit For
                        End If
                    Next j
                End If
            End If
        Next i
        
        debugMsg = debugMsg & "  Processed " & processedOps & " operations from evaluation" & vbCrLf
        debugMsg = debugMsg & "  Matched " & matchedOps & " operations in HeatMap" & vbCrLf
        debugMsg = debugMsg & "  Updated " & updatedCount & " status dots" & vbCrLf & vbCrLf
    End If
    
    ' Step 7: Summary
    Application.StatusBar = "Step 7/7: Complete"
    debugMsg = debugMsg & "✓ Step 7: Update complete!" & vbCrLf & vbCrLf
    
    Dim elapsedTime As Double
    elapsedTime = Timer - startTime
    debugMsg = debugMsg & "=== SUMMARY ===" & vbCrLf
    debugMsg = debugMsg & "Operations updated: " & updatedCount & vbCrLf
    debugMsg = debugMsg & "Time taken: " & Format(elapsedTime, "0.00") & " seconds" & vbCrLf
    
    Application.ScreenUpdating = True
    Application.StatusBar = False
    
    If updatedCount > 0 Then
        MsgBox debugMsg, vbInformation, "HeatMap Update Successful"
    Else
        debugMsg = debugMsg & vbCrLf & "⚠ WARNING: No operations were updated!" & vbCrLf & vbCrLf
        debugMsg = debugMsg & "POSSIBLE CAUSES:" & vbCrLf
        debugMsg = debugMsg & "1. Operation codes in HeatMap don't match Evaluation Results" & vbCrLf
        debugMsg = debugMsg & "2. No status values found in Final Status column" & vbCrLf
        debugMsg = debugMsg & "3. All statuses are 'N/A' or blank" & vbCrLf
        MsgBox debugMsg, vbExclamation, "No Updates Made"
    End If
    
    Exit Sub

ErrorHandler:
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Error occurred: " & Err.Description & vbCrLf & vbCrLf & _
           "Error Number: " & Err.Number & vbCrLf & _
           "Line: " & Erl, vbCritical, "Error"
End Sub

' Helper function to create the update button
Sub CreateUpdateButtonWithDebug()
    Dim ws As Worksheet
    Dim btn As Button
    Dim btnExists As Boolean
    
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("HeatMap Sheet")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "HeatMap Sheet not found! Please create or rename the sheet to 'HeatMap Sheet'", _
               vbCritical, "Sheet Not Found"
        Exit Sub
    End If
    
    ' Check if button already exists
    btnExists = False
    On Error Resume Next
    Set btn = ws.Buttons("UpdateHeatMapButton")
    If Not btn Is Nothing Then btnExists = True
    On Error GoTo 0
    
    If btnExists Then
        If MsgBox("Update button already exists. Replace it?", vbYesNo + vbQuestion, "Button Exists") = vbNo Then
            Exit Sub
        End If
        btn.Delete
    End If
    
    ' Create new button
    Set btn = ws.Buttons.Add(100, 10, 200, 30)
    With btn
        .Name = "UpdateHeatMapButton"
        .Text = "Update HeatMap Status (Debug)"
        .OnAction = "UpdateHeatMapStatusWithDebug"
    End With
    
    MsgBox "Button created successfully on HeatMap Sheet!" & vbCrLf & vbCrLf & _
           "Click the button to update status and see detailed debug information.", _
           vbInformation, "Button Created"
End Sub
