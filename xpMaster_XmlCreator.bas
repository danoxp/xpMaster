Attribute VB_Name = "XmlCreator"
Option Explicit

Sub Tester()

    Dim XML As Object, rt As Object, nd As Object, i As Long, n As Long
    
    Set XML = EmptyDocument()
    
    Set rt = CreateXmlElement(XML, "Root", , Array("name", "dano", "color", "blue"))
''    Set rt = CreateXmlElement(XML, "Root")
    XML.appendChild rt
    
    For i = 1 To 3
        Set nd = CreateXmlElement(XML, "config", "CFG" & i, Array("type", "Typ" & i))
        rt.appendChild nd
        For n = 1 To 4
            nd.appendChild _
                 CreateXmlElement(XML, "item", "ITM" & n, _
                                      Array("name", "It's a Test " & n))
        Next n
    Next i
''    Debug.Print XML.XML
    Debug.Print PrettyPrintXML(XML.XML)
End Sub


' ### everything below here is a utility method ###

'Utility method: create and return an element, with
'   optional value and attributes
Public Function CreateXmlElement(doc As Object, elementName As String, _
                            Optional elementValue As String, _
                            Optional attributesArray As Variant = Empty, _
                            Optional parentEl As Object) As Object
    '// passed in as Array(attr1Name, attr1Value, attr2Name, attr2Value,...)
    Dim e
    Dim i As Long
    Dim o As Object
    
    Set e = doc.CreateNode(1, elementName, "")  '// create empty node 'elementName'
    
    '// if have attributes, loop and add
    If Not IsEmpty(attributesArray) Then
        For i = 0 To UBound(attributesArray) Step 2
            Set o = doc.CreateAttribute(attributesArray(i))
            o.Value = attributesArray(i + 1)
            e.Attributes.setNamedItem o
        Next i
    End If
    
    'any element content to add?
    If VBA.Len(elementValue) > 0 Then
        Set o = doc.createTextNode(elementValue)
        e.appendChild o
    End If
    
    If Not parentEl Is Nothing Then parentEl.appendChild e
    
    Set CreateXmlElement = e
End Function

'create and return an empty xml doc
Public Function EmptyDocument() As Object
    Dim XML
    Set XML = CreateObject("MSXML2.DOMDocument")
    XML.LoadXML ""
    XML.appendChild XML.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'") '// 'UTF-16'
    Set EmptyDocument = XML
End Function

'https://stackoverflow.com/questions/1118576/how-can-i-pretty-print-xml-source-using-vb6-and-msxml
Public Function PrettyPrintXML(s As String) As String
  Dim Writer As Object 'New MXXMLWriter60
''  Dim Reader As Object 'New SAXXMLReader60

''  Debug.Print s
  Set Writer = CreateObject("MSXML2.MXXMLWriter.6.0")
  With Writer
        .indent = True
        .standalone = False
        .omitXMLDeclaration = False
        .Encoding = "UTF-8" '// UTF-16 default
        .byteOrderMark = False
  End With
  
  With CreateObject("MSXML2.SAXXMLReader.6.0")
      Set .contentHandler = Writer
      Set .dtdHandler = Writer
      Set .errorHandler = Writer
      Call .putProperty("http://xml.org/sax/properties/declaration-handler", _
              Writer)
      Call .putProperty("http://xml.org/sax/properties/lexical-handler", _
              Writer)
      Call .parse(s)
  End With

  PrettyPrintXML = Writer.output
End Function

'// https://stackoverflow.com/a/6406372/4363840
Private Sub FormatDocToFile(ByVal doc As MSXML2.DOMDocument60, _
                            ByVal FileName As String)
    'Reformats the DOMDocument "Doc" into an ADODB.Stream
    'and writes it to the specified file.
    '
    'Note the UTF-8 output never gets a BOM.  If we want one we
    'have to write it here explicitly after opening the Stream.
    Dim rdrDom As MSXML2.SAXXMLReader60
    Dim stmFormatted As ADODB.Stream
    Dim wtrFormatted As MSXML2.MXXMLWriter60
    Dim o As MSXML2.DOMDocument60
    Set stmFormatted = New ADODB.Stream
    With stmFormatted
        .Open
        .Type = adTypeBinary
        Set wtrFormatted = New MSXML2.MXXMLWriter60
        With wtrFormatted
            .omitXMLDeclaration = False
            .standalone = True
            .byteOrderMark = False 'If not set (even to False) then
                                   '.encoding is ignored.
            .Encoding = "utf-8"    'Even if .byteOrderMark = True
                                   'UTF-8 never gets a BOM.
            .indent = True
            .output = stmFormatted
            Set rdrDom = New MSXML2.SAXXMLReader60
            With rdrDom
                Set .contentHandler = wtrFormatted
                Set .dtdHandler = wtrFormatted
                Set .errorHandler = wtrFormatted
                .putProperty "http://xml.org/sax/properties/lexical-handler", _
                             wtrFormatted
                .putProperty "http://xml.org/sax/properties/declaration-handler", _
                             wtrFormatted
                .parse doc
            End With
        End With
        .SaveToFile FileName
        .Close
    End With
End Sub


