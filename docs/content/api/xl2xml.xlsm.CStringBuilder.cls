'####Module : CStringBuilder
'>#####Type : Class Module
'>#####Description : Efficient string contatenator
'***
Option Explicit

Private Const MODULE As String = "CStringBuilder"

Private m_lngMax As Long        'maxsize of the memory buffer, should be a good estimate, larger than required
Private m_lngCounter As Long    'position in memory block
Private m_strResult As String

Property Get CurrentBufferMaxSize() As Long
    CurrentBufferMaxSize = m_lngMax
End Property

'***
Private Sub Class_Initialize()
'***
'>Description : VBA default class initilization
'>Parameters  : none
'>Returns : the class
'>Dependencies: none
'>Notes: none
'>Usage :
'```
'Dim oMySB as New CStringBuilder
'Debug.Assert(Not(oMySB Is Nothing))
'```
    Const METHODNAME As String = "Class_Initialize"
    m_lngCounter = 1
End Sub

'***
Public Sub Initialise(ByVal l_lngSize As Long)
'***
'>Description : performs any initialisation necessary for the class
'>Parameters  :
'> > 1. l_lngSize : the initial maximum size of the buffer.  Better to make it longer than needed.
'>
'>Returns : none
'>Dependencies: none
'>Notes: meant to be called from code if there is any initialisation that should be passed from the caller
'>Usage :
'```
'Dim oMySB as New CStringBuilder
'oMySB.Initialise 500
'Debug.Assert(oMySB.CurrentBufferMaxSize = 500)
'```
    Const METHODNAME As String = "Initialise"
    m_strResult = String(l_lngSize, " ")
End Sub


'***
Public Sub Append(ByRef p_strToAppend As String)
'***
'>Description : appends to the end of the string
'>Parameters  :
'> > 1. p_strToAppend : the string to append to the builder
'>
'>Returns : none
'>Dependencies: none
'>Notes: none
'>Usage :
'```
'Dim oMySB as New CStringBuilder
'oMySB.Initialise 500
'oMySB.Append "hello"
'Debug.Assert(oMySB.ToString() = "hello")
'```
    Const METHODNAME As String = "Append"
    Mid$(m_strResult, m_lngCounter, Len(p_strToAppend)) = p_strToAppend
    m_lngCounter = m_lngCounter + Len(p_strToAppend)
End Sub

'***
Public Function ToString() As String
'***
'>Description : writes the result from the builder
'>Parameters  : none
'>Returns : the string that has been built up.
'>Dependencies: none
'>Notes: none
'>Usage :
'```
'Dim oMySB as New CStringBuilder
'oMySB.Initialise 500
'oMySB.Append "hello"
'Debug.Assert(oMySB.ToString() = "hello")
'```
    Const METHODNAME As String = "ToString"
    ToString = Left$(m_strResult, m_lngCounter)
End Function


