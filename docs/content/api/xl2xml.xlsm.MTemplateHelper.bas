'####Module : MTemplateHelper
'#####Type : Module
'#####Description : Contains methods to handle some of the inconsistencies and situations in the template
'***
Option Explicit
Option Private Module

Private Const MODULENAME As String = "MTemplateHelper"

'***
Public Function GetChildFromParentChild(ByRef p_strIn As String) As String
'***
'>Description : Returns the child given the parent/child path string
'>Parameters  :
'> > 1. p_strIn : the hierachy string to parse
'>
'>Returns : an empty string in the event of an error, the substring following the slash otherwise
'>Dependencies: none
'>Notes: Doesn't matter what direction the slash is in
'>Usage :
'```
'Dim s as string : s = "iati-activities\iati-activity"
'Debug.Assert(GetChildFromParentChild(s) = "iati-activity")
'```
    Const METHODNAME As String = "GetChildFromParentChild"
    On Error GoTo ErrorTrap
    
    Dim sResult As String
    If (Not (InStr(p_strIn, "/") > 0) And Not (InStr(p_strIn, "\") > 0)) Then
        Err.Raise vbError + 2001, METHODNAME, "Unable to find parent child delimiter."
    Else
        If (InStr(p_strIn, "/") > 0) Then
            sResult = Right$(p_strIn, Len(p_strIn) - InStr(p_strIn, "/"))
        End If
        If (InStr(p_strIn, "\") > 0) Then
            sResult = Right$(p_strIn, Len(p_strIn) - InStr(p_strIn, "\"))
        End If
    End If
    
    GoTo PrivateExitPoint
    
ErrorTrap:
    sResult = vbNullString
PrivateExitPoint:
    
    GetChildFromParentChild = sResult
End Function

'***
Public Function GetParentFromParentChild(ByRef p_strIn As String) As String
'***
'>Description : Returns the parent given the parent/child path string
'>Parameters  :
'> > 1. p_strIn : the hierachy string to parse
'>
'>Returns : an empty string in the event of an error, the substring preceeding the slash otherwise
'>Dependencies: none
'>Notes: Doesn't matter what direction the slash is in
'>Usage :
'```
'Dim s as string : s = "iati-activities\iati-activity"
'Debug.Assert(GetParentFromParentChild(s) = "iati-activities")
'```
    Const METHODNAME As String = "GetParentFromParentChild"
    On Error GoTo ErrorTrap
    
    Dim sResult As String
    If (Not (InStr(p_strIn, "/") > 0) And Not (InStr(p_strIn, "\") > 0)) Then
        Err.Raise vbError + 2001, METHODNAME, "Unable to find parent child separator."
    Else
        If (InStr(p_strIn, "/") > 0) Then
            sResult = Left$(p_strIn, InStr(p_strIn, "/") - 1)
        End If
        If (InStr(p_strIn, "\") > 0) Then
            sResult = Left$(p_strIn, InStr(p_strIn, "\") - 1)
        End If
    End If
    
    GoTo PrivateExitPoint
    
ErrorTrap:
    sResult = vbNullString
PrivateExitPoint:
    
    GetParentFromParentChild = sResult
End Function


'***
Public Function GetStartingDataColumn(ByRef p_strIn As String) As Integer
'***
'>Description : Gets the first column that has data in it, not meta information
'>Parameters  :
'> > 1. p_strIn : the name of the worksheet to search on
'>
'>Returns : the column number of the column where the data starts
'>Dependencies: none
'>Notes:
'>Usage :
'```
'Dim i as integer : i = GetStartingDataColumn("Activity Dates")
'Debug.Assert(i = 6)
'```
    Const METHODNAME As String = "GetStartingDataColumn"
    Dim i As Integer
    
    With ThisWorkbook.Worksheets(p_strIn)
        For i = 3 To .Range("ZZ17").End(xlToLeft).Column
            If (.Range(MExcelHelper.ColumnNumbertoLetter(i) & "17").Value <> "N/A") And (Trim(.Range(MExcelHelper.ColumnNumbertoLetter(i) & "17").Value) <> "") Then
                Exit For
            End If
        Next i
    End With
        
    GetStartingDataColumn = i

End Function

'***
Public Function IsAttribute(ByRef p_strIn As String) As Boolean
'***
'>Description : Tests whether the string starts with the @ attribute indicator
'>Parameters  :
'> > 1. p_strIn : the string to test
'>
'>Returns : true in the event it does, false otherwise
'>Dependencies: none
'>Notes:
'>Usage :
'```
'Dim s as integer : s = "@type=string"
'Debug.Assert(true = IsAttribute(s))
'```
    Const METHODNAME As String = "IsAttribute"
    If (InStr(p_strIn, "@") > 0) Then
        IsAttribute = True
    Else
        IsAttribute = False
    End If
End Function

'***
Public Function GetAttribute(ByRef p_strIn As String) As String
'***
'>Description : Parses something like ```iati-activity@date``` to return ```date```
'>Parameters  :
'> > 1. p_strIn : the string to parse
'>
'>Returns : an empty string in the event the attribute character is missing, the attribute otherwise.
'>Dependencies: none
'>Notes : none
'>Usage :
'```
'Dim s as integer : s = "iati-activity@type"
'Debug.Assert("type" = GetAttribute(s))
'```
    Const METHODNAME As String = "GetAttribute"
    On Error GoTo ErrorTrap
    
    Dim sResult As String
    If (Not (InStr(p_strIn, "@") > 0)) Then
        Err.Raise vbError + 2001, METHODNAME, "Unable to determine attribute name"
    Else
        sResult = Right$(p_strIn, Len(p_strIn) - InStr(p_strIn, "@"))
    End If
    
    GoTo PrivateExitPoint
    
ErrorTrap:
    sResult = vbNullString
PrivateExitPoint:
    
    GetAttribute = sResult
End Function

'***
Public Function GetAttributeParent(ByRef p_strIn As String) As String
'***
'>Description : Parses something like ```iati-activity@date``` to return ```iati-activity```
'>Parameters  :
'> > 1. p_strIn : the string to parse
'>
'>Returns : an empty string in the event the attribute character is missing, the parent element otherwise.
'>Dependencies: none
'>Notes : none
'>Usage :
'```
'Dim s as integer : s = "iati-activity@type"
'Debug.Assert("iati-activity" = GetAttributeParent(s))
'```
    Const METHODNAME As String = "GetAttributeParent"
    On Error GoTo ErrorTrap
    
    Dim sResult As String
    If (Not (InStr(p_strIn, "@") > 0)) Then
        Err.Raise vbError + 2001, METHODNAME, "Unable to determine attribute name"
    Else
        sResult = Left$(p_strIn, InStr(p_strIn, "@") - 1)
    End If
    
    GoTo PrivateExitPoint
    
ErrorTrap:
    sResult = vbNullString
PrivateExitPoint:
    
    GetAttributeParent = sResult
End Function

'***
Public Function GetCurrentElement(ByRef p_strIn As String) As String
'***
'>Description : Parses something like ```iati-activity@date``` to return ```iati-activity```
'>Parameters  :
'> > 1. p_strIn : the string to parse
'>
'>Returns : an empty string in the event the attribute character is missing, the current element otherwise.
'>Dependencies: none
'>Notes : Can parse 'elementName', 'elementName\childElementName' or 'elementName@attribute' to return 'elementName'
'>Usage :
'```
'Dim s as integer : s = "iati-activity@type"
'Debug.Assert("iati-activity" = GetCurrentElement(s))
'```
    Const METHODNAME As String = "GetCurrentElement"
    Dim sResult As String
    If (sResult = vbNullString) Then sResult = GetAttributeParent(p_strIn)
    If (sResult = vbNullString) Then sResult = GetParentFromParentChild(p_strIn)
    If (sResult = vbNullString) Then sResult = p_strIn
       
    GetCurrentElement = sResult
End Function

'***
Public Function IsRawElement(ByRef p_strIn As String) As Boolean
'***
'>Description :  Whether the string contains "@" or "\"
'>Parameters  :
'> > 1. p_strIn : the string to test
'>
'>Returns : false in the event the string contains complex characters ("@" or "\"), true otherwise
'>Dependencies: none
'>Notes : none
'>Usage :
'```
'Dim s as integer : s = "iati-activity"
'Debug.Assert(true = IsRawElement(s))
'```
    Const METHODNAME As String = "IsRawElement"
    If (InStr(p_strIn, "@") > 0) Then
        IsRawElement = False
    ElseIf (InStr(p_strIn, "\") > 0) Then
        IsRawElement = False
    Else
        IsRawElement = True
    End If
End Function

'***
Public Function IsComplexElement(ByRef p_strIn As String) As Boolean
'***
'>Description :  Whether the string contains "\"
'>Parameters  :
'> > 1. p_strIn : the string to test
'>
'>Returns : false in the event the string contains a slash character, true otherwise
'>Dependencies: none
'>Notes : none
'>Usage :
'```
'Dim s as integer : s = "iati-activity\activity-date"
'Debug.Assert(true = IsComplexElement(s))
'```
    Const METHODNAME As String = "IsComplexElement"
    If (InStr(p_strIn, "\") > 0) Then
        IsComplexElement = True
    Else
        IsComplexElement = False
    End If
End Function


