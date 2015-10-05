Attribute VB_Name = "MXL2XMLGlobals"
'####Module : MGlobals
'#####Type : Module
'#####Description : Holds the global references
'***

Option Explicit
Option Private Module

Private Const MODULENAME As String = "MGlobals"

' Global variables
Public Const gVersion As String = "0.9.0.0"
Public gMyXDoc As New MSXML2.DOMDocument60
Public gFileSaveName As String

