Option Explicit

' ============================================================================
' Module: UpdateHeatMapStatus
' Purpose: Update HeatMap Sheet Status column with evaluation results
' Author: GitHub Copilot
' Date: 2025-12-08
' ============================================================================

' Column constants for Evaluation Results sheet
Private Const EVAL_OP_CODE_COLUMN As Long = 1
Private Const EVAL_OPERATION_COLUMN As Long = 2
Private Const EVAL_FINAL_STATUS_COLUMN As Long = 12

' Column constants for HeatMap Sheet
Private Const HEATMAP_OP_CODE_COLUMN As Long = 1
Private Const HEATMAP_OPERATION_COLUMN As Long = 2
Private Const HEATMAP_STATUS_COLUMN As Long = 18  ' Column R

' Row constants
Private Const EVAL_DATA_START_ROW As Long = 2
Private Const HEATMAP_DATA_START_ROW As Long = 4

' ============================================================================
' Main Subroutine - Called by the "Update HeatMap Status" button
' ============================================================================
Public Sub UpdateHeatMapStatus()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim wsEval As Worksheet
    Dim wsHeatmap As Worksheet
    Dim evalData As Object  ' Dictionary to store evaluation data
    Dim lastRow As Long
    Dim row As Long
    Dim opCode As Variant
    Dim operation As String
    Dim finalStatus As String
    Dim updatesCount As Long
    Dim noMatchCount As Long
    
    ' Disable screen updating for better performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Set wb = ThisWorkbook
    
    ' Check if required sheets exist
    On Error Resume Next
    Set wsEval = wb.Worksheets("Evaluation Results")
    Set wsHeatmap = wb.Worksheets("HeatMap Sheet")
    On Error GoTo ErrorHandler
    
    If wsEval Is Nothing Then
        MsgBox "Error: 'Evaluation Results' sheet not found!", vbCritical, "Update Status"
        GoTo CleanUp
    End If
    
    If wsHeatmap Is Nothing Then
        MsgBox "Error: 'HeatMap Sheet' not found!", vbCritical, "Update Status"
        GoTo CleanUp
    End If
    
    ' Create dictionary for evaluation data
    Set evalData = CreateObject("Scripting.Dictionary")
    
    ' Read Evaluation Results
    lastRow = wsEval.Cells(wsEval.Rows.Count, EVAL_OP_CODE_COLUMN).End(xlUp).row
    
    For row = EVAL_DATA_START_ROW To lastRow
        opCode = wsEval.Cells(row, EVAL_OP_CODE_COLUMN).Value
        operation = wsEval.Cells(row, EVAL_OPERATION_COLUMN).Value
        finalStatus = wsEval.Cells(row, EVAL_FINAL_STATUS_COLUMN).Value
        
        ' Only process rows with valid op codes
        If Not IsEmpty(opCode) And IsNumeric(opCode) Then
            Dim opCodeStr As String
            opCodeStr = CStr(CLng(opCode))
            
            ' Add status to dictionary (store multiple statuses per op code)
            If evalData.Exists(opCodeStr) Then
                evalData(opCodeStr) = evalData(opCodeStr) & "|" & finalStatus
            Else
                evalData.Add opCodeStr, finalStatus
            End If
        End If
    Next row
    
    ' Update HeatMap Sheet
    lastRow = wsHeatmap.Cells(wsHeatmap.Rows.Count, HEATMAP_OP_CODE_COLUMN).End(xlUp).row
    updatesCount = 0
    noMatchCount = 0
    
    For row = HEATMAP_DATA_START_ROW To lastRow
        opCode = wsHeatmap.Cells(row, HEATMAP_OP_CODE_COLUMN).Value
        operation = wsHeatmap.Cells(row, HEATMAP_OPERATION_COLUMN).Value
        
        If Not IsEmpty(opCode) And Not IsEmpty(operation) And IsNumeric(opCode) Then
            Dim opCodeStr2 As String
            opCodeStr2 = CStr(CLng(opCode))
            
            ' Find matching evaluations
            If evalData.Exists(opCodeStr2) Then
                Dim statusList As String
                Dim worstStatus As String
                
                statusList = evalData(opCodeStr2)
                worstStatus = DetermineWorstStatus(statusList)
                
                ' Update the cell
                wsHeatmap.Cells(row, HEATMAP_STATUS_COLUMN).Value = worstStatus
                updatesCount = updatesCount + 1
            Else
                noMatchCount = noMatchCount + 1
            End If
        End If
    Next row
    
    ' Show summary message
    MsgBox "Update Complete!" & vbCrLf & vbCrLf & _
           "Updated: " & updatesCount & " operations" & vbCrLf & _
           "No match: " & noMatchCount & " operations", _
           vbInformation, "Update Status"
    
CleanUp:
    ' Re-enable screen updating
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Error: " & Err.Description, vbCritical, "Update Status"
End Sub

' ============================================================================
' Function: DetermineWorstStatus
' Purpose: Determine the worst status from a pipe-separated list
' Parameters: statusList - Pipe-separated string of statuses
' Returns: The worst status (RED > YELLOW > GREEN)
' ============================================================================
Private Function DetermineWorstStatus(statusList As String) As String
    Dim statuses() As String
    Dim i As Long
    Dim currentStatus As String
    Dim worstPriority As Long
    Dim currentPriority As Long
    Dim worstStatus As String
    
    ' Split the status list
    statuses = Split(statusList, "|")
    
    ' Initialize with worst possible priority
    worstPriority = 999
    worstStatus = ""
    
    ' Find the worst status
    For i = LBound(statuses) To UBound(statuses)
        currentStatus = Trim(statuses(i))
        currentPriority = GetStatusPriority(currentStatus)
        
        ' Lower priority number = worse status
        If currentPriority < worstPriority Then
            worstPriority = currentPriority
            worstStatus = currentStatus
        End If
    Next i
    
    ' Return empty string if no valid status found
    If worstPriority = 999 Then
        DetermineWorstStatus = ""
    Else
        DetermineWorstStatus = worstStatus
    End If
End Function

' ============================================================================
' Function: GetStatusPriority
' Purpose: Get priority for status values (lower = worse)
' Parameters: status - Status string to evaluate
' Returns: Priority number (0=worst, 3=neutral)
' ============================================================================
Private Function GetStatusPriority(status As String) As Long
    Dim statusUpper As String
    
    ' Handle empty or N/A values
    If status = "" Or IsEmpty(status) Then
        GetStatusPriority = 3  ' Neutral
        Exit Function
    End If
    
    statusUpper = UCase(Trim(status))
    
    Select Case statusUpper
        Case "RED"
            GetStatusPriority = 0  ' Worst
        Case "YELLOW"
            GetStatusPriority = 1  ' Medium
        Case "GREEN"
            GetStatusPriority = 2  ' Good
        Case "N/A"
            GetStatusPriority = 3  ' Neutral
        Case Else
            GetStatusPriority = 3  ' Unknown/N/A
    End Select
End Function

' ============================================================================
' Subroutine: ClearHeatMapStatus
' Purpose: Clear all status values from HeatMap Sheet (optional utility)
' ============================================================================
Public Sub ClearHeatMapStatus()
    On Error GoTo ErrorHandler
    
    Dim wb As Workbook
    Dim wsHeatmap As Worksheet
    Dim lastRow As Long
    Dim row As Long
    
    Set wb = ThisWorkbook
    Set wsHeatmap = wb.Worksheets("HeatMap Sheet")
    
    If MsgBox("Are you sure you want to clear all status values?", _
              vbYesNo + vbQuestion, "Clear Status") = vbNo Then
        Exit Sub
    End If
    
    Application.ScreenUpdating = False
    
    lastRow = wsHeatmap.Cells(wsHeatmap.Rows.Count, HEATMAP_OP_CODE_COLUMN).End(xlUp).row
    
    For row = HEATMAP_DATA_START_ROW To lastRow
        wsHeatmap.Cells(row, HEATMAP_STATUS_COLUMN).ClearContents
    Next row
    
    Application.ScreenUpdating = True
    
    MsgBox "Status column cleared successfully!", vbInformation, "Clear Status"
    Exit Sub
    
ErrorHandler:
    Application.ScreenUpdating = True
    MsgBox "Error: " & Err.Description, vbCritical, "Clear Status"
End Sub
