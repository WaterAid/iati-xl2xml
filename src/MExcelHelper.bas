Attribute VB_Name = "MExcelHelper"
'####Module : MExcelHelper
'#####Type : Module
'#####Description : Functions to work with the grid and other aspects of Excel
'***
Option Explicit
Option Private Module

Private Const MODULENAME As String = "MExcelHelper"

'***
Public Function ColumnNumbertoLetter(ByVal intColumn As Integer) As String
'***
'>Description : Converts the column number to the column letter
'>Parameters  :
'> > 1. intColumn : the number to convert to a column letter
'>
'>Returns : the column letter given the number
'>Dependencies: none
'>Notes: none
'>Usage :
'```
'Dim s as string
's = ColumnNumbertoLetter(4)
'Debug.Assert(s = "d")
'```
    Const METHODNAME As String = "ColumnNumbertoLetter"
    Dim n As Integer
    Dim C As Byte
    Dim s As String

    n = intColumn
    Do
        C = ((n - 1) Mod 26)
        s = Chr(C + 65) & s
        n = (n - C) \ 26
    Loop While n > 0
    ColumnNumbertoLetter = s
End Function


'***
Public Sub ResetAppSettings()
'***
'>Description : Resets the application environment to the usual settings
'>Parameters : none
'>Returns : none
'>Dependencies: none
'>Notes: none
'>Usage :
'```
'ResetAppSettings
'```
    Const METHODNAME As String = "ResetAppSettings"
    With Application
        .Calculation = xlCalculationAutomatic
        .ScreenUpdating = True
        .DisplayAlerts = True
        .Cursor = xlDefault
        .StatusBar = vbNullString
    End With

End Sub
