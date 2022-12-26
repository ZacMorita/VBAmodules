Attribute VB_Name = "Module1"
Option Explicit

'This Class is written to work with both "Early Binding" and "Late Binding" _
    of the "Microsoft XML, v6.0" Object Library. _
    If compilation fails with message: _
    "Compile Error: User-Defined Type Not Defined" then, in the _
    Visual Basic Editor Go to Tools >> References... >> Then check the box _
    next to the mentioned library. Or, if Late Binding is required, _
    manually change "#Const EarlyBinding = True" to "#Const EarlyBinding = False" _
    Early Binding is exponentially faster and less error prone than Late Binding.

#Const EarlyBinding = True '<--- This is manually adjufted for Early and Late Binding

Const XmlPath As String = "C:\EXAMPLES\EXAMPLE\Example.xml"
Const NewXmlPath As String = "C:\EXAMPLES\EXAMPLE\Example2.xml"

Public Sub ManipulateWord()
    'IMPORTANT NOTE: _
        If compilation fails with message "Compile Error: User-Defined Type Not Defined" _
        Please see "EarlyBinding" Comment at the top of this Class Module.
    'EarlyBinding is a pre-compile constant at the top of this module
    #If EarlyBinding Then 'This is a pre-compile conditional
        Dim doc As MSXML2.DOMDocument60
        Dim nodes As IXMLDOMNodeList
        Dim node As IXMLDOMNode
        Set doc = New MSXML2.DOMDocument60
    #Else
        Dim doc As Object
        Dim nodes As Object
        Dim node As Object
        Set doc = CreateObject("MSXML2.DOMDocument.6.0")
    #End If

    doc.Load XmlPath

    '"//person" is an XPath expression"
    Set nodes = doc.SelectNodes("//person")

    For Each node In nodes
        '"primary_identifier" and "username" are both XPath expressions"
        node.SelectSingleNode("primary_identifier").Text = _
            node.SelectSingleNode("username").Text
    Next node

    doc.Save NewXmlPath

    Set doc = Nothing

End Sub

'=======================================
'VVVVVVVVVVV Example.xml VVVVVVVVVVVVVVV
'=======================================
'<root>
'    <person>
'        <primary_identifier>john.doe@username.com</primary_identifier>
'        <username>jdoe123</username>
'    </person>
'    <person>
'        <primary_identifier>smith.doe@username.com</primary_identifier>
'        <username>sdoe123</username>
'    </person>
'    <person>
'        <primary_identifier>jain.doe@username.com</primary_identifier>
'        <username>jdoe124</username>
'    </person>
'</root>
