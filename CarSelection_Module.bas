Attribute VB_Name = "CarSelection"
Option Explicit

' Module: CarSelection
' Purpose: Manage dynamic car selection for Target and Tested vehicles
' Location: Dropdowns placed after column V (in columns W and X)

' ===============================================
' CONSTANTS FOR DROPDOWN LOCATIONS
' ===============================================
Const TARGET_CAR_DROPDOWN_COL As String = "W"
Const TESTED_CAR_DROPDOWN_COL As String = "X"
Const DROPDOWN_ROW As Long = 1
Const CAR_DATA_START_COL As Long = 8  ' Column H where car data starts
Const CAR_NAME_ROW As Long = 1  ' Row where car names are located

' ===============================================
' INITIALIZE DROPDOWNS ON SHEET1
' ===============================================
Public Sub InitializeCarSelectionDropdowns()
    Dim wsSheet1 As Worksheet
    Dim carNames As String
    Dim lastCol As Long
    Dim i As Long
    Dim cellAddress As String
    
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    
    ' Find last column with data in row 1 (car names)
    lastCol = wsSheet1.Cells(CAR_NAME_ROW, wsSheet1.Columns.Count).End(xlToLeft).Column
    
    ' Build comma-separated list of car names from columns H onwards
    carNames = ""
    For i = CAR_DATA_START_COL To lastCol
        If Trim(CStr(wsSheet1.Cells(CAR_NAME_ROW, i).Value)) <> "" Then
            If carNames <> "" Then carNames = carNames & ","
            carNames = carNames & Trim(CStr(wsSheet1.Cells(CAR_NAME_ROW, i).Value))
        End If
    Next i
    
    ' Setup Target Car Dropdown (Column W)
    With wsSheet1.Range(TARGET_CAR_DROPDOWN_COL & DROPDOWN_ROW)
        .ClearContents
        .Value = "Select Target Car"
        .Font.Bold = True
        .Interior.Color = RGB(255, 242, 204)  ' Light yellow
        .HorizontalAlignment = xlCenter
        
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:=carNames
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End With
    
    ' Setup Tested Car Dropdown (Column X)
    With wsSheet1.Range(TESTED_CAR_DROPDOWN_COL & DROPDOWN_ROW)
        .ClearContents
        .Value = "Select Tested Car"
        .Font.Bold = True
        .Interior.Color = RGB(204, 255, 204)  ' Light green
        .HorizontalAlignment = xlCenter
        
        With .Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
                 Formula1:=carNames
            .IgnoreBlank = True
            .InCellDropdown = True
        End With
    End With
    
    ' Add labels in row above if space available
    wsSheet1.Range(TARGET_CAR_DROPDOWN_COL & "2").Value = "Target:"
    wsSheet1.Range(TARGET_CAR_DROPDOWN_COL & "2").Font.Bold = True
    wsSheet1.Range(TESTED_CAR_DROPDOWN_COL & "2").Value = "Tested:"
    wsSheet1.Range(TESTED_CAR_DROPDOWN_COL & "2").Font.Bold = True
    
    MsgBox "Car selection dropdowns initialized in columns W and X!" & vbCrLf & _
           "Please select Target and Tested cars before running evaluation.", vbInformation
End Sub

' ===============================================
' GET SELECTED CAR NAMES
' ===============================================
Public Function GetTargetCarName() As String
    Dim wsSheet1 As Worksheet
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    
    GetTargetCarName = Trim(CStr(wsSheet1.Range(TARGET_CAR_DROPDOWN_COL & DROPDOWN_ROW).Value))
    If GetTargetCarName = "Select Target Car" Then GetTargetCarName = ""
End Function

Public Function GetTestedCarName() As String
    Dim wsSheet1 As Worksheet
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    
    GetTestedCarName = Trim(CStr(wsSheet1.Range(TESTED_CAR_DROPDOWN_COL & DROPDOWN_ROW).Value))
    If GetTestedCarName = "Select Tested Car" Then GetTestedCarName = ""
End Function

' ===============================================
' FIND COLUMN INDEX FOR CAR NAME
' ===============================================
Public Function FindCarColumn(wsSheet As Worksheet, carName As String) As Long
    Dim lastCol As Long
    Dim i As Long
    
    FindCarColumn = 0
    If Trim(carName) = "" Then Exit Function
    
    lastCol = wsSheet.Cells(CAR_NAME_ROW, wsSheet.Columns.Count).End(xlToLeft).Column
    
    For i = CAR_DATA_START_COL To lastCol
        If UCase(Trim(CStr(wsSheet.Cells(CAR_NAME_ROW, i).Value))) = UCase(Trim(carName)) Then
            FindCarColumn = i
            Exit Function
        End If
    Next i
End Function

' ===============================================
' VALIDATE CAR SELECTIONS
' ===============================================
Public Function ValidateCarSelections() As Boolean
    Dim targetCar As String
    Dim testedCar As String
    
    targetCar = GetTargetCarName()
    testedCar = GetTestedCarName()
    
    If targetCar = "" Or testedCar = "" Then
        MsgBox "Please select both Target and Tested cars from the dropdowns in columns W and X.", _
               vbExclamation, "Car Selection Required"
        ValidateCarSelections = False
        Exit Function
    End If
    
    If targetCar = testedCar Then
        Dim result As VbMsgBoxResult
        result = MsgBox("Target and Tested cars are the same (" & targetCar & ")." & vbCrLf & _
                       "Do you want to continue?", vbQuestion + vbYesNo, "Same Car Selected")
        If result = vbNo Then
            ValidateCarSelections = False
            Exit Function
        End If
    End If
    
    ValidateCarSelections = True
End Function

' ===============================================
' GET COLUMN OFFSETS FOR SELECTED CARS
' ===============================================
' Returns the column indices for Target and Tested cars
' This will be used to dynamically read data based on selections
Public Sub GetSelectedCarColumns(ByRef targetCol As Long, ByRef testedCol As Long)
    Dim wsSheet1 As Worksheet
    Dim targetCar As String
    Dim testedCar As String
    
    Set wsSheet1 = ThisWorkbook.Sheets("Sheet1")
    
    targetCar = GetTargetCarName()
    testedCar = GetTestedCarName()
    
    targetCol = FindCarColumn(wsSheet1, targetCar)
    testedCol = FindCarColumn(wsSheet1, testedCar)
    
    If targetCol = 0 Then
        MsgBox "Could not find Target car: " & targetCar, vbCritical
    End If
    
    If testedCol = 0 Then
        MsgBox "Could not find Tested car: " & testedCar, vbCritical
    End If
End Sub
