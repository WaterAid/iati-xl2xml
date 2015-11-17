'####Module : MXMLHelper
'#####Type : Module
'#####Description : Functions to work with aspects of XML
'***
Option Explicit
Option Private Module

Private Const MODULENAME As String = "MXMLHelper"

'***
Public Function IsEmptyElement(ByRef p_objXElement As MSXML2.IXMLDOMElement) As Boolean
'***
'>Description : Tests if an element is named but otherwise empty
'>Parameters :
'> > 1. p_objXElement : the element to test
'>
'>Returns : true if the element is named and empty, false otherwise
'>Dependencies: none
'>Usage :
'```
'Dim oMyElement as new MSXML2.IXMLDOMElement
'Debug.Assert(false = IsEmptyElement(oMyElement))
'```
    Const METHODNAME As String = "IsEmptyElement"
    IsEmptyElement = False
    If (p_objXElement.Attributes.Length = 0 And Not (p_objXElement.HasChildNodes()) And p_objXElement.nodeTypedValue = vbNullString) Then IsEmptyElement = True
    
End Function

'***
Public Function ConvertToSpecialCharacter(ByRef p_strIn As String) As String
'***
'>Description : converts the characters that need xml escaping
'>Parameters :
'> > 1. p_strIn : the raw string to perform the replacements
'>
'>Returns : the xml escaped string
'>Dependencies: none
'>Notes : makes the following conversions: & to &amp;, < to &lt;, > to &gt;, " to &quot;,' to &apos;,
'>Usage :
'```
'Dim sRawString as String : s = "<hello>"
'Debug.Assert(s  = "&lt;hello&gt;")
'```
    Const METHODNAME As String = "ConvertToSpecialCharacter"
    ConvertToSpecialCharacter = Trim(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(Replace(p_strIn, Chr(34), "&quot;"), "&", "&amp;"), "'", "&apos;"), "<", "&lt;"), ">", "&gt;"), vbCrLf, ""), vbCr, ""), vbLf, ""), vbBack, ""), vbFormFeed, ""), vbNewLine, ""), vbNullChar, ""), vbNullString, ""), vbTab, ""), vbVerticalTab, ""))

End Function

