Attribute VB_Name = "exportCode"
Option Explicit

'// #INCLUDE: [Microsoft Visual Basic for Applications Extensibility]
'// #INCLUDE: [MSXML2]
'// #INCLUDE: XmlCreator.bas Module

Private m As thisModule
Private Type thisModule
    s As String
    fldr As String
    gitfldr As String
    dic As Scripting.Dictionary
End Type

Public Sub ExportAllVBAcode()
    '// Exports all code in open Workbooks and installed Addins
    '// including Worksheet XML and Workbook VBA code
    '// sheet XML files for data to rebuild sheets with formatting and formulas
    Dim isUserAddInsChanged As Boolean
    Dim i As Long
    Dim s As String
    
    '// target directory:  %Appdata%\Git\name
    m.gitfldr = Environ("APPDATA") & "\Git\"
    If VBA.Len(VBA.Dir(m.gitfldr, vbDirectory)) = 0 Then VBA.MkDir m.gitfldr
    
    '// 'Trust' VBE object model, then turn off when finished
    If Not isVBEPermissionsOn Then MsgBox "cannot export without VBE permissions, exit", vbInformation: Exit Sub
    
    Debug.Print vbLf & "Export Directory: ["; m.gitfldr; "]"; vbLf; "---------------------"
    
    isUserAddInsChanged = isSelectedAddIns '// skip Addins menu at end if no changes
    
    Set m.dic = New Dictionary      '// keep track of git Folder names to prevent duplicates
    
    '// Export all open WorkBooks
    Debug.Print vbLf & "Excel.Workbooks:"; Excel.Workbooks.Count; vbLf; "---------------------"
    For i = 1 To Excel.Workbooks.Count
        With Workbooks(i)
''            Debug.Print , .Name, IIf(.HasVBProject, .VBProject.Name, vbTab), IIf(.Saved, vbTab, "not-saved")
            Debug.Print vbLf; i; "'"; .Name; "'",
        
            Select Case True
            Case Not .Name Like "*.xl*"
                Debug.Print , .Name; " skipped"
            Case Not .Saved
                Debug.Print .Name; ": Not Saved - skipped"
            Case Not .HasVBProject  '// does this ever happen? All wb have VBProject?
                Debug.Print "No VBProject"
                If VBA.MsgBox(.Name & " has No VBProject.  Save Workbook Xml?", vbYesNo, .FullName) = vbYes Then
                    exportWorkbook Workbooks(i)
                End If
            Case .VBProject.Protection = vbext_pp_locked
                Debug.Print .Name; ": Protected - skipped"
                Debug.Assert False
            Case Else
                exportWorkbook Workbooks(i)
            End Select
        End With
    Next i

    '// Export all loaded AddIns
    Debug.Print vbLf; "Excel.AddIns:"; Excel.AddIns.Count; vbLf; "---------------------"

    For i = 1 To Excel.AddIns.Count
        With AddIns(i)
''            Debug.Print vbLf; i; "'"; .Name; "'"; ,
            
            Select Case True
            Case Not .Installed
                s = s & vbLf & vbTab & i & " '" & .Name & "'"
                
''                Debug.Print "Not Installed"
            Case Not Workbooks(.Name).Saved
                Debug.Print vbLf; i; "'"; .Name; "'"; ,
                Debug.Print "Not Saved - skipped"
                MsgBox .Name & " is not-saved, skipped"
            Case Else
                Debug.Print vbLf; i; "'"; .Name; "'"; ,
                exportWorkbook Workbooks(.Name)
            End Select
        End With
    Next i
    
    Debug.Print "Not Installed:"; s
    
    '// List all Installed COMAddIns in immediate window only
    Debug.Print vbLf; "Application.COMAddIns:"; Application.COMAddIns.Count; vbLf; "---------------------"
    For i = 1 To Application.COMAddIns.Count
        With Application.COMAddIns(i)
            Debug.Print vbLf; i; .progID; vbLf, "["; .Description; "]"; vbLf, .GUID
        End With
    Next i
    
    If isUserAddInsChanged Then Debug.Print isSelectedAddIns    '// only open AddInsSelection if changed
    If Not isVBEPermissionsOff Then MsgBox "VBE permissions are on, dangerous", vbCritical
End Sub

Private Sub exportWorkbook(Wb As Excel.Workbook)
    Dim XML As Object   '// Document
    Dim rt As Object    '// root for nodes to add
    Dim nd As Object    '// node
    
    With Wb.VBProject
        
        '// Git subfolder [Begin] name and check it:
        m.s = .Name
        If m.s = "VBAProject" Then          '// replace generic 'VBAProject' with filename prefix
            m.s = Replace(.BuildFileName, ".DLL", vbNullString)
            m.s = VBA.Mid(m.s, VBA.InStrRev(m.s, "\") + 1)
        End If
        
        '// avoid overwriting an already exported a Workbook when two projects have the same name
        If m.dic.Exists(m.s) Then
            Dim s As String
            s = m.dic(m.s) & vbLf & vbLf & .BuildFileName & vbLf & "overwrite 1st with 2nd?"
            If MsgBox(s, vbYesNo, "Duplicate VBProject: " & m.s) = vbNo Then
                '// use filename prefix instead
                m.s = Replace(.BuildFileName, ".DLL", vbNullString)
                m.s = VBA.Mid(m.s, VBA.InStrRev(m.s, "\") + 1)
            End If
        End If
        
        m.dic(m.s) = .BuildFileName
        
        
        m.fldr = m.gitfldr & m.s & "\"
        If VBA.Len(VBA.Dir(m.fldr, vbDirectory)) = 0 Then VBA.MkDir m.fldr
        Debug.Print vbLf; , ; "["; m.s; "]"
        '// Git subfolder [End]
        
        Set XML = XmlCreator.EmptyDocument()
        '// rt is ExcelFile
        Set rt = CreateXmlElement(XML, "ExcelFile", , Array("IsAddin", IIf(Wb.IsAddin, "True", "False"), "Name", Wb.Name), XML)
''        If Wb.IsAddin Then rt.setAttribute "IsAddin", "True"
        '// nd is WorkBook
        Set nd = CreateXmlElement(XML, "WorkBook", , , rt)
        Call CreateXmlElement(XML, "ProjectName", .Name, , nd)
        Call CreateXmlElement(XML, "FullName", Wb.FullName, , nd)
        Call CreateXmlElement(XML, "Path", Wb.Path, , nd)
        Call CreateXmlElement(XML, "FileName", Wb.Name, , nd)
''        Call CreateXmlElement(XML, "IsAddin", Wb.IsAddin, , nd)
        Call CreateXmlElement(XML, "Author", Wb.Author, , nd)
        Call CreateXmlElement(XML, "Description", .Description, , nd)
    End With
    
    addSheets2Xml Wb, XML, rt   '// WorkBook, XmlDocument, ExcelFile node
    
    addVBProject Wb.VBProject, XML, rt
    
    addReferences2Xml Wb.VBProject, XML, rt
    
''    CreateObject("scripting.filesystemobject").CreateTextFile(m.fldr & m.s & ".xml").Write PrettyPrintXML(XML.XML)
    saveTextToFile PrettyPrintXML(XML.XML), m.fldr & m.s & ".xml", "utf-8"
    
    Debug.Print , m.s & ".xml"  '' & vbTab & m.fldr
End Sub

Private Sub addVBProject(project As VBProject, doc As Object, parente As Object)
    Dim rt As Object
    Dim nd As Object
    Dim i As Long
    Dim D
    
    Set rt = CreateXmlElement(doc, "VBComponents", , , parente)
    For i = 1 To project.VBComponents.Count: With project.VBComponents(i)
        Do
            If .Type = vbext_ct_Document And .CodeModule.CountOfLines < 3 Then Exit Do
''            Set nd = CreateXmlElement(doc, "VBComponent", , Array("ID", i, "Type", VBA.Choose(.Type, "StdModule", "ClassModule", "MSForm"), "Name", .Name), rt)
            Set nd = CreateXmlElement(doc, "VBComponent", , Array("id", "vbc" & i), rt)
                
            Select Case .Type
            Case vbext_ct_Document
                .Export m.fldr & m.s & "_" & .Name & ".vb"
                Debug.Print , m.s & "_" & .Name & ".vb"
                Call CreateXmlElement(doc, "CodeFile", .Name & ".vb", , nd)
                nd.setAttribute "Type", "Document"
            Case vbext_ct_StdModule
                .Export m.fldr & m.s & "_" & .Name & ".bas"
                Debug.Print , m.s & "_" & .Name & ".bas"
                Call CreateXmlElement(doc, "CodeFile", .Name & ".bas", , nd)
                nd.setAttribute "Type", "StdModule"
            Case vbext_ct_ClassModule
                .Export m.fldr & m.s & "_" & .Name & ".cls"
                Debug.Print , m.s & "_" & .Name & ".cls"
                Call CreateXmlElement(doc, "CodeFile", .Name & ".cls", , nd)
                nd.setAttribute "Type", "ClassModule"
            Case vbext_ct_MSForm
                .Export m.fldr & m.s & "_" & .Name & ".frm"
                Debug.Print , m.s & "_" & .Name & ".frm"
                Call CreateXmlElement(doc, "CodeFile", .Name & ".frm", , nd)
                nd.setAttribute "Type", "MSForm"
            Case Else
                '// .Type = vbext_ct_ActiveXDesigner
                Debug.Assert False
            End Select
            
            nd.setAttribute "Name", .Name
            Call CreateXmlElement(doc, "CountOfDeclarationLines", .CodeModule.CountOfDeclarationLines, , nd)
            Call CreateXmlElement(doc, "CountOfLines", .CodeModule.CountOfLines, , nd)
        Loop Until True
    End With: Next i

End Sub

Private Sub addSheets2Xml(Wb As Workbook, doc As Object, parente As Object)
    Dim i As Long
    Dim nd As Object
    Dim rt As Object ', rrt As Object
    Dim filenm As String
    Dim sxml As String
    Dim c As Range
    Dim N As Long
    
    Set rt = XmlCreator.CreateXmlElement(doc, "Sheets", , Array("Count", Wb.Sheets.Count), parente)
''    Set fso = CreateObject("scripting.filesystemobject")

    For i = 1 To Wb.Sheets.Count: With Wb.Sheets(i)
        Set nd = CreateXmlElement(doc, .CodeName, , Array("id", "sh" & i, "Type", VBA.TypeName(Wb.Sheets(i)), "Name", .Name), rt)
        Call CreateXmlElement(doc, "Name", .Name, , nd)
        Call CreateXmlElement(doc, "CodeName", .CodeName, , nd)
        If .Visible <> XlSheetVisibility.xlSheetVisible Then
            Call CreateXmlElement(doc, "Visible", IIf(.Visible = xlSheetHidden, "Hidden", "VeryHidden"), , nd)
        End If

        Select Case VBA.TypeName(Wb.Sheets(i))
        
        Case "Worksheet"
            Do
                '// skip blank sheets
                If VBA.IsEmpty(.UsedRange) Then
                    Call CreateXmlElement(doc, "CellsCount", "0", , nd)
                    Exit Do
                End If
                
                '// UsedRange, UsedCells
                Call CreateXmlElement(doc, "UsedRange", .UsedRange.AddressLocal, , nd)
                Call CreateXmlElement(doc, "CellsCount", .UsedRange.Cells.Count, , nd)
                N = 0
                For Each c In .UsedRange.Cells
                    If VBA.IsEmpty(c) Then N = N + 1
                Next c
                Call CreateXmlElement(doc, "CellsEmpty", VBA.CStr(N), , nd)
                
                '// write out WorkSheet Xml file of worksheet as excel import format
                filenm = m.fldr & m.s & "_" & .Name & ".xml"
                Call CreateXmlElement(doc, "XmlFilename", filenm, , nd) '// add filename to Xml
                
                '// get XML string
                '// xlRangeValueXMLSpreadsheet  - Excel w formats, formulas, and names
                '// xlRangeValueMSPersistXML    - Recordset format as XML
                '//   - sometimes, XMLSpreadsheet must be called before MSPersistXML for this to work!
                '// include Cells(1) in output Range to get full sheet
                sxml = .Range(.Cells(1), .UsedRange.Cells(.UsedRange.Cells.Count)).Value(xlRangeValueXMLSpreadsheet)
                saveTextToFile sxml, filenm, "utf-8"
                Debug.Print , filenm
            Loop Until True
        
        Case "Chart"
            '// save chart png and add file to Xml
            Call CreateXmlElement(doc, "image", .Name & ".png", , nd)
            .Export FileName:=m.fldr & m.s & "_" & .Name & ".png", FilterName:="png"
            Debug.Print , m.s & "_" & .Name & ".png"
        
        Case Else
            Debug.Assert False
        
        End Select
        
        '// Shapes added to Xml [TODO] is there any way to save Shapes as png?
        addShapes2Xml Wb.Sheets(i), doc, nd
    End With: Next i

''    Set fso = Nothing
End Sub

Public Sub saveTextToFile(content, filePath, charset)
    '// omegastripes JSON2XML.bas
    With CreateObject("ADODB.Stream")
        .Type = 2 ' adTypeText
        .Open
        .charset = charset
        .WriteText content
        .Position = 0
        .Type = 1 ' TypeBinary
        .SaveToFile filePath, 2
        .Close
    End With
End Sub

Private Sub addShapes2Xml(sh As Object, doc As Object, parentt As Object)
    Dim rt As Object
    Dim nd As Object
    Dim i As Long ', j As Long
''    Dim rrt As Object
''    Dim sp As Excel.Shape
    
    If sh.Shapes.Count = 0 Then Exit Sub
    
    Debug.Print , "- [Shapes:"; sh.Shapes.Count & "]"
    Set rt = XmlCreator.CreateXmlElement(doc, "Shapes", , Array("Count", sh.Shapes.Count), parentt)
    
    For i = 1 To sh.Shapes.Count
        With sh.Shapes(i)
    ''    Set sp = sh.Shapes(i)
            Set nd = CreateXmlElement(doc, shapeTypeName(.Type) & "-" & i, , Array("ZOrder", .ZOrderPosition, "id", "shp" & .ID, "Type", shapeTypeName(.Type), "Name", .Name), rt)
            Call CreateXmlElement(doc, "ZOrderPosition", .ZOrderPosition, , nd)
            Call CreateXmlElement(doc, "ID", .ID, , nd)
            Call CreateXmlElement(doc, "Name", .Name, , nd)
            Call CreateXmlElement(doc, "Type", shapeTypeName(.Type), , nd)
            Call CreateXmlElement(doc, "Dimensions", "{" & .Left & ", " & .Top & ", " & .Width & ", " & .Height & "}", _
                Array("Left", .Left, "Top", .Top, "Width", .Width, "Height", .Height), nd)
            If Len(.AlternativeText) > 0 Then _
                Call CreateXmlElement(doc, "AlternativeText", VBA.Replace(Replace(.AlternativeText, vbCr, "\r"), vbLf, "\n"), , nd)
            If TypeName(sh) = "Worksheet" Then _
                Call CreateXmlElement(doc, "Range", "[" & .TopLeftCell.Address & ":" & .BottomRightCell.Address & "]", _
                Array("TopLeftCell", .TopLeftCell.Address, "BottomRightCell", .BottomRightCell.Address), nd)
            Debug.Print , "  "; i; shapeTypeName(.Type), "[" & .Name & "]" ': Stop
            
            Select Case .Type   '// MsoShapeType
            Case msoChart ': Stop
                Call CreateXmlElement(doc, "ChartName", .Chart.Name, , nd)
                If .Chart.HasTitle Then Call CreateXmlElement(doc, "ChartTitle", .Chart.ChartTitle.Caption, , nd)
                Call CreateXmlElement(doc, "ChartType", .Chart.ChartType, , nd)
                Call CreateXmlElement(doc, "ChartStyle", .Chart.ChartStyle, , nd)
                Call CreateXmlElement(doc, "image", .Chart.Name & ".png", , nd)
                .Chart.Export FileName:=m.fldr & m.s & "_" & .Chart.Name & ".png", FilterName:="png"
                Debug.Print , , "["; m.s & "_" & .Chart.Name & ".png]"
            Case msoComment ': Stop
                '// comments are included in SheetXml file
            Case msoTextBox    '// add Caption text
                Call CreateXmlElement(doc, "Caption", .DrawingObject.Caption, , nd) '// same as .DrawingObject.Text
            Case msoAutoShape ': Stop
    ''            Call CreateXmlElement(doc, "ChartName", .Chart.Name, , nd)
            Case msoPicture
                '// AlternativeText already added
            Case msoSmartArt
                '// info is in GroupItems.Items(j).TextFrame2.TextRange.Text
            Case msoEmbeddedOLEObject ': Stop
                Call CreateXmlElement(doc, "ProgID", .OLEFormat.progID, , nd) '// 'Paint.Picture'
            Case msoOLEControlObject
                '// .Name 'Control 1'
            Case Else
                Debug.Assert False
''            Case msoCallout: Stop
''            Case msoFreeform: Stop
''            Case msoGroup: Stop
''            Case msoFormControl: Stop
''            Case msoLine: Stop
''            Case msoLinkedOLEObject: Stop
''            Case msoLinkedPicture: Stop
''            Case msoPlaceholder: Stop
''            Case msoTextEffect: Stop
''            Case msoMedia: Stop
''            Case msoScriptAnchor: Stop
''            Case msoTable: Stop
''            Case msoCanvas: Stop
''            Case msoDiagram: Stop
''            Case msoInk: Stop
''            Case msoInkComment: Stop
''            Case msoShapeTypeMixed: Stop
            End Select
            
        End With
    Next i
End Sub

Function shapeTypeName(N As MsoShapeType) As String
    Dim v
    
    v = VBA.Choose(N, "AutoShape", "Callout", "Chart", "Comment", "Freeform", "Group", _
        "EmbeddedOLEObject", "FormControl", "Line", "LinkedOLEObject", "LinkedPicture", _
        "OLEControlObject", "Picture", "Placeholder", "TextEffect", "Media", "TextBox", _
        "ScriptAnchor", "Table", "Canvas", "Diagram", "Ink", "InkComment", "SmartArt")
    If Not IsNull(v) Then
        shapeTypeName = v
    Else
        shapeTypeName = "ShapeTypeMixed"
    End If
End Function

Private Sub addReferences2Xml(pj As VBIDE.VBProject, doc As Object, parente As Object)
    Dim i As Long
    Dim nd As Object
    Dim ret As Object
    
    Set ret = XmlCreator.CreateXmlElement(doc, "References", , , parente)
    
    For i = 1 To pj.References.Count
        With pj.References(i)
            Set nd = CreateXmlElement(doc, "Reference", , Array("id", "ref" & i, "Type", .Type, "BuiltIn", IIf(.BuiltIn, "True", "False"), "Name", .Name), ret)
            Call CreateXmlElement(doc, "Description", .Description, , nd)
            If VBA.Len(.Description) > 0 Then Call CreateXmlElement(doc, "FullPath", .FullPath, , nd)
            Call CreateXmlElement(doc, "Version", .Major & "." & .Minor, , nd)
''            Call CreateXmlElement(doc, "BuiltIn", .BuiltIn, , nd)
            Call CreateXmlElement(doc, "GUID", .GUID, , nd)
            If .IsBroken Then
                MsgBox .Name & " has a broken reference to: " & .Name, vbCritical
                Call CreateXmlElement(doc, "isBroken", .IsBroken, , nd)
            End If
        End With
    Next i
End Sub

Private Function isSelectedAddIns() As Boolean  '// did user change installed Addins?
    Dim i As Long
    Dim N As Long
    
    For i = 1 To Excel.AddIns.Count
        If AddIns(i).Installed Then N = N + i
    Next i
    
    Debug.Print "Select Addins to Export Code"
    Application.Dialogs(xlDialogAddinManager).Show  '// .Dialogs(321).Show
    
    For i = 1 To Excel.AddIns.Count '// check to see if Addins were selected/deselected
        If AddIns(i).Installed Then N = N - i
    Next i
    isSelectedAddIns = (N <> 0)
    
End Function

Private Function isVBEPermissionsOn() As Boolean
    On Error Resume Next
        If Not Application.VBE.VBProjects.Count > 0 Then
            Debug.Print vbLf; "enable 'Trust Access' to 'VBE Project Object'"
            Application.CommandBars.ExecuteMso "MacroSecurity"  '// turn off macroSecurity
        '// Application.CommandBars.FindControl(ID:=3627).Execute  '//same thing
        Else
            Debug.Print vbLf; "VBE Project Ojbect' already exposed w 'Trust Access' (dangerous)"
        End If
    isVBEPermissionsOn = IsNumeric(Application.VBE.VBProjects.Count)
End Function

Private Function isVBEPermissionsOff() As Boolean
    Debug.Print vbLf; "disable 'Trust Access' to 'VBA Project Object' for safety"
    Application.CommandBars.ExecuteMso "MacroSecurity"
    On Error Resume Next
    Debug.Assert IsNumeric(Application.VBE.VBProjects.Count)
    isVBEPermissionsOff = (Err.Number = 1004)
End Function
