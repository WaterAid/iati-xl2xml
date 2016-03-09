'####Module : MGenerateXml
'#####Type : Module
'#####Description : Functions to work with the grid and other aspects of Excel
'***
Option Explicit

Private Const MODULENAME As String = "MXmlGenerator"

'***
Public Sub GenerateXml()
'***
'>Description : Looks through current work book and generates XML to satisfy IATI schema.
'>Parameters  : none
'>Returns : none
'>Dependencies:
'> > 1. gFileSaveName
'> > 2. MExcelHelper
'> > 3. gMyXDoc
'>
'>Notes: Looks for user input to save file, starts at Activity Main and works through each sheet populating and writing elements, attributes and data. File written in Unicode.
'>Usage : This is the top level method, it is not called in code but rather invoked from the macro window.
    Const METHODNAME As String = "GenerateXml"
    Dim dbResult As Double
    Dim i As Integer
    Dim strLastUpdate As String

    On Error GoTo ErrorTrap
    
    ' get the file
    gFileSaveName = Application.GetSaveAsFilename(ThisWorkbook.Sheets("Prerequisites").Range("G10").Value, "XML File,*.xml", , "Save file as")
    If gFileSaveName = "False" Then MExcelHelper.ResetAppSettings
    ' get the last update
    strLastUpdate = Format(Now(), "YYYY-MM-DD") & "T" & Format(Now(), "HH:MM:SS") & "Z"
    
    ' create the root element
    Dim oMyNewNode As MSXML2.IXMLDOMElement
    Set oMyNewNode = gMyXDoc.createNode(MSXML2.NODE_ELEMENT, "iati-activities", "http://www.w3.org/XML/1998/namespace")
    oMyNewNode.setAttribute "version", ThisWorkbook.Sheets("iati-activities").Range("C21").Value
    oMyNewNode.setAttribute "xmlns:xsi", "http://www.w3.org/2001/XMLSchema-instance"
    oMyNewNode.setAttribute "generated-datetime", strLastUpdate
    Set gMyXDoc.DocumentElement = oMyNewNode
    gMyXDoc.Save gFileSaveName   ' this method lets you save as you go and the practice is to call it every activity
    
    ' now take a walk through Activity Main Information looking for elements
    ElementScan "Activity Main Information", oMyNewNode, vbNullString
        
 
    
    dbResult = i

    GoTo PrivateExitPoint
    
ErrorTrap:

    Debug.Print Err.Description
    dbResult = -1
    
PrivateExitPoint:

    MExcelHelper.ResetAppSettings
End Sub


'***
Public Function ElementScan(ByRef p_strWorksheetName As String, _
    ByRef p_objMyParentElement As MSXML2.IXMLDOMElement, ByRef p_strKey As String) As Double
'***
'>Description : Does the element scan for each row in activity main information it scans deep for the complex elements (i.e. into other worksheets) then across for the simple elements and attributes
' Then goes to the next line/ activity saving as it goes.
'>Parameters  :
'> > 1. p_strWorksheetName : the worksheet to scan
'> > 2. p_objMyParentElement : the parent xml element at the current point in the scan
'> > 3. p_strKey : the activity key
'>
'>Returns : -1 in the event of error, 0 otherwise.  In practice this is not captured and progress is determined by how many activities make it into the outbound file.
'>Dependencies:
'> > 1. MTemplateHelper
'> > 2. MExcelHelper
'> > 3. MXmlHelper
'> > 4. gMyXDoc
'>
'>Notes:
'>Usage :
'```
'ElementScan "Activity Main Information", oMyNewNode, vbNullString
'```
    Const METHODNAME As String = "ElementScan"
    Dim sParent As String
    Dim sRootElement As String
    Dim strKey As String
    Dim iColumnIterator As Integer
    Dim l As Long
    Dim i As Integer

    sRootElement = MTemplateHelper.GetChildFromParentChild(CStr(ThisWorkbook.Worksheets(p_strWorksheetName).Range("C1").Value))
    sParent = MTemplateHelper.GetParentFromParentChild(CStr(ThisWorkbook.Worksheets(p_strWorksheetName).Range("C1").Value))
    
    ' only continue if context is correct i.e. you are looking at the right parent
    If (p_objMyParentElement.nodeName = sParent) Then
        
        With ThisWorkbook.Worksheets(p_strWorksheetName)
            .Activate       ' may not be necessary but jumping around a bit
            ' find the key by dropping down column C from row 21
            For l = 21 To .Range("C65000").End(xlUp).Row
                strKey = .Range("C" & l).Value
                               
                ' only continue if your keys match in other words if you have found a row with the key you hold
                If (strKey = p_strKey Or p_strKey = vbNullString) Then  ' Or is a special case for the first row
                    Application.StatusBar = strKey & ": Processing..."
                    Dim oFirstChild As MSXML2.IXMLDOMElement
                    Set oFirstChild = gMyXDoc.createElement(MTemplateHelper.GetCurrentElement(sRootElement))
                    
                    ' find the data by scanning across the current row from the first column that doesn't say meta
                    iColumnIterator = MTemplateHelper.GetStartingDataColumn(p_strWorksheetName)
                    
                    ' set the head element
                    Dim oHeadElement As MSXML2.IXMLDOMElement
                    Set oHeadElement = gMyXDoc.createElement(MTemplateHelper.GetCurrentElement(.Range(MExcelHelper.ColumnNumbertoLetter(iColumnIterator) & "20")))
                    Dim oCurrentElement As MSXML2.IXMLDOMElement
                    
                    For i = iColumnIterator To .Range("ZZ17").End(xlToLeft).Column
                        
                        ' set the current element
                        Set oCurrentElement = gMyXDoc.createElement(MTemplateHelper.GetCurrentElement(.Range(MExcelHelper.ColumnNumbertoLetter(i) & "20")))
                                            
                        ' test for new current element
                        If (Not (oHeadElement.nodeName = oCurrentElement.nodeName)) Then
                            
                            ' you have moved on to a new current element, you want to recurse on the
                            ' head at this point in case there are children of it
                            Dim wk As Worksheet
                            For Each wk In ThisWorkbook.Worksheets
                                If (Not (wk.Name = p_strWorksheetName)) Then
                                    ElementScan wk.Name, oHeadElement, strKey
                                End If
                            Next wk
                                                        
                            ' attach the head element to the parent because you are finished processing it
                            ' special case is if it's the first time then you replace
                            If (MXMLHelper.IsEmptyElement(oFirstChild)) Then
                                Set oFirstChild = oHeadElement
                            Else
                                oFirstChild.appendChild oHeadElement
                            End If
                            
                            ' reset the head to the
                            Set oHeadElement = oCurrentElement
                        
                        End If
                        
                        ' process this node
                        If (MTemplateHelper.IsAttribute(.Range(MExcelHelper.ColumnNumbertoLetter(i) & "20"))) Then
                            ' add an attribute to the current element
                            Dim sAttributeName As String
                            sAttributeName = MTemplateHelper.GetAttribute(.Range(MExcelHelper.ColumnNumbertoLetter(i) & "20"))
                            oHeadElement.setAttribute sAttributeName, .Range(MExcelHelper.ColumnNumbertoLetter(i) & CStr(l))
                        End If
                            
                        If (MTemplateHelper.IsRawElement(.Range(MExcelHelper.ColumnNumbertoLetter(i) & "20"))) Then
                            ' add the value to the current element
                            oHeadElement.nodeTypedValue = MXMLHelper.ConvertToSpecialCharacter(.Range(MExcelHelper.ColumnNumbertoLetter(i) & CStr(l)).Value2)
                        End If
                            
                        If (MTemplateHelper.IsComplexElement(.Range(MExcelHelper.ColumnNumbertoLetter(i) & "20"))) Then
                            ' there is a deliberate limitation here, you can only have one child in an inline complex element
                            ' in other words only: contact-info\organisation, not contact-info\organisation\department
                            ' if you want to put that much detail in you must use extra worksheets
                                
                            oHeadElement.appendChild gMyXDoc.createElement(MTemplateHelper.GetChildFromParentChild(.Range(MExcelHelper.ColumnNumbertoLetter(i) & "20")))
                            oHeadElement.LastChild.nodeTypedValue = MXMLHelper.ConvertToSpecialCharacter(.Range(MExcelHelper.ColumnNumbertoLetter(i) & CStr(l)).Value2)
                        End If
                            
                        ' test if this is the final column to process
                        If (i = .Range("ZZ17").End(xlToLeft).Column) Then
                            ' it's the final column so add what you've processed so far
                            ' special case is if it's the first time then you replace
                            If (MXMLHelper.IsEmptyElement(oFirstChild)) Then
                                Set oFirstChild = oHeadElement
                            Else
                                oFirstChild.appendChild oHeadElement
                            End If
                        End If
                        
                    Next i
                    
                    ' attach the first child to the parent
                    p_objMyParentElement.appendChild oFirstChild
                    ' do an interim save if you have just completed an activity
                    If (p_objMyParentElement.nodeName = "iati-activities") Then
                        
                        Application.StatusBar = strKey & ": Complete!"
                        Debug.Print oFirstChild.XML
                        
                        gMyXDoc.FirstChild.appendChild oFirstChild
                        gMyXDoc.Save gFileSaveName
                        
                    End If
                      
                End If
                
            Next l
        End With
    End If

End Function


