Attribute VB_Name = "exportCode"
Option Explicit

'// #INCLUDE: [Microsoft Visual Basic for Applications Extensibility]
'// #INCLUDE: [MSXML2]

Private m As thisModule
Private Type thisModule
    s As String
    fldr As String
    gitfldr As String
End Type

Public Sub ExportAllVBAcode()
    '// Exports all code in open Workbooks and installed Addins
    '// including Worksheet XML and Workbook VBA code
    '// sheet XML files for data to rebuild sheets with formatting and formulas
    Dim userAddinsSelected As Boolean
    Dim i As Long
    
    '// target directory:  %Appdata%\Git\name
    m.gitfldr = Environ("APPDATA") & "\Git\"
    If VBA.Len(VBA.Dir(m.gitfldr, vbDirectory)) = 0 Then VBA.MkDir m.gitfldr
    
    '// 'Trust' VBE object model, then turn off when finished
    If Not isVBEPermissionsOn Then MsgBox "cannot export without VBE permissions, exit", vbInformation: Exit Sub
    
    userAddinsSelected = isSelectedAddins '// skip Addins menu at end if no changes
    
    Debug.Print vbLf & "Workbooks:"
    For i = 1 To Excel.Workbooks.Count
        With Workbooks(i)
            Debug.Print , IIf(.HasVBProject, .VBProject.Name, vbTab), IIf(.Saved, vbTab, "not-saved"), .Name
        
            Select Case True
            Case Not .HasVBProject
            Case Not .Saved: MsgBox .Name & " is not saved, skipped"
            Case .VBProject.Protection = vbext_pp_locked: MsgBox .Name & " is protected, skipped"
            Case Else
                exportWorkbook Workbooks(i)
            End Select
        End With
    Next i

    Debug.Print "Addins:"
    For i = 1 To Excel.AddIns.Count
        With AddIns(i)
            Select Case True
            Case Not .Installed: Debug.Print "  -off-", , , .Name
            Case Not Workbooks(.Name).Saved: MsgBox .Name & " is not-saved, skipped"
            Case Else
                exportWorkbook Workbooks(.Name)
            End Select
        End With
    Next i
    
    Debug.Print "COM Addins"
    For i = 1 To Application.COMAddIns.Count
        With Application.COMAddIns(i)
        Debug.Print .GUID, .progID, .Description
        End With
    Next i
    
    If userAddinsSelected Then Debug.Print isSelectedAddins
    If Not isVBEPermissionsOff Then MsgBox "VBE permissions are on, dangerous", vbCritical
End Sub

Private Sub exportWorkbook(wb As Excel.Workbook)
    Dim XML As Object
    Dim rt As Object
    Dim nd As Object
    
    With wb.VBProject
    
        '// Git subfolder name and check it:
        m.s = .Name
        If m.s = "VBAProject" Then                      '// use filename instead of generic VBAProject
            m.s = Replace(.BuildFileName, ".DLL", vbNullString)
            m.s = VBA.Mid(m.s, VBA.InStrRev(m.s, "\") + 1)
        End If
        m.fldr = m.gitfldr & m.s & "\"
        If VBA.Len(VBA.Dir(m.fldr, vbDirectory)) = 0 Then VBA.MkDir m.fldr
        '// Git subfolder [End]
        
        Set XML = XmlCreator.EmptyDocument()
        Set rt = CreateXmlElement(XML, "ExcelFile", , Array("Name", wb.Name), XML)
        If wb.IsAddin Then rt.setAttribute "IsAddin", "True"
        Set nd = CreateXmlElement(XML, "Meta", , , rt)
        Call CreateXmlElement(XML, "ProjectName", .Name, , nd)
        Call CreateXmlElement(XML, "FileName", wb.Name, , nd)
        Call CreateXmlElement(XML, "Path", wb.Path, , nd)
        Call CreateXmlElement(XML, "IsAddin", wb.IsAddin, , nd)
        Call CreateXmlElement(XML, "Author", wb.Author, , nd)
        Call CreateXmlElement(XML, "Description", .Description, , nd)
    End With
    
    AddSheets2Xml wb, XML, rt
    
    ExportVBProject wb.VBProject, XML, rt
    
    AddReferences2Xml wb.VBProject, XML, rt
    
    With CreateObject("scripting.filesystemobject")
        .CreateTextFile(m.fldr & m.s & ".xml").Write PrettyPrintXML(XML.XML)
    End With
    
''    Debug.Print PrettyPrintXML(XML.XML)
    Debug.Print vbTab & m.s & ".xml" '' & vbTab & m.fldr

End Sub

Private Sub ExportVBProject(project As VBProject, doc As Object, parente As Object)
    Dim rt As Object
    Dim nd As Object
    Dim i As Long
    Dim s As String
    
    Set rt = CreateXmlElement(doc, "VBComponents", , , parente)
    For i = 1 To project.VBComponents.Count
        With project.VBComponents(i)
            
            Set nd = CreateXmlElement(doc, .Name, , Array("Id", i), rt)
            If .CodeModule.CountOfLines > 2 Then
                
                Select Case .Type
                Case vbext_ct_StdModule
                    .Export m.fldr & m.s & "_" & .Name & ".bas"
                    Call CreateXmlElement(doc, "CodeFile", .Name & ".bas", , nd)
                    nd.setAttribute "Type", "StdModule"
                Case vbext_ct_Document
                    .Export m.fldr & m.s & "_" & .Name & ".vb"
                    Call CreateXmlElement(doc, "CodeFile", .Name & ".vb", , nd)
                    nd.setAttribute "Type", "Document"
                Case vbext_ct_ClassModule
                    .Export m.fldr & m.s & "_" & .Name & ".cls"
                    Call CreateXmlElement(doc, "CodeFile", .Name & ".cls", , nd)
                    nd.setAttribute "Type", "ClassModule"
                Case vbext_ct_MSForm
                    .Export m.fldr & m.s & "_" & .Name & ".frm"
                    Call CreateXmlElement(doc, "CodeFile", .Name & ".frm", , nd)
                    nd.setAttribute "Type", "MSForm"
                Case Else       '// .Type = vbext_ct_ActiveXDesigner
                    Debug.Assert False
                End Select
                
                Call CreateXmlElement(doc, "CountOfDeclarationLines", .CodeModule.CountOfDeclarationLines, , nd)
                Call CreateXmlElement(doc, "CountOfLines", .CodeModule.CountOfLines, , nd)
            End If
        End With
    Next i

End Sub

Private Sub AddSheets2Xml(wb As Workbook, doc As Object, parente As Object)
    Dim fso As Object
    Dim i As Long
    Dim nd As Object
    Dim rt As Object
    
    Set rt = CreateXmlElement(doc, "Sheets", , , parente)
    Set fso = CreateObject("scripting.filesystemobject")

    For i = 1 To wb.Sheets.Count
        With wb.Sheets(i)
            Set nd = CreateXmlElement(doc, .CodeName, , Array("Id", i, "Type", VBA.TypeName(wb.Sheets(i))), rt)
            If .CodeName <> .Name Then Call CreateXmlElement(doc, "Name", .Name, , nd)
       
            Select Case True
            Case .Type <> xlWorksheet       '// skip charts
            Case VBA.IsEmpty(.UsedRange)    '// skip blank sheets
            Case Else
                Call CreateXmlElement(doc, "UsedRange", .UsedRange.AddressLocal, , nd)
                Call CreateXmlElement(doc, "XmlFilename", m.s & "_" & .CodeName & ".xml", , nd)
                fso.CreateTextFile(m.fldr & m.s & "_" & .CodeName & ".xml").Write .UsedRange.Value(xlRangeValueXMLSpreadsheet)
            End Select
        End With
    Next i

    Set fso = Nothing
    End Sub
    
Private Sub AddReferences2Xml(pj As VBIDE.VBProject, doc As Object, parente As Object)
    Dim i As Long
    Dim nd As Object
    Dim ret As Object
    
    Set ret = XmlCreator.CreateXmlElement(doc, "References", , , parente)
    
    For i = 1 To pj.References.Count
        With pj.References(i)
            Set nd = CreateXmlElement(doc, .Name, , , ret)
            Call CreateXmlElement(doc, "Description", .Description, , nd)
            Call CreateXmlElement(doc, "Version", .Major & "." & .Minor, , nd)
            Call CreateXmlElement(doc, "BuiltIn", .BuiltIn, , nd)
            Call CreateXmlElement(doc, "GUID", .GUID, , nd)
            If VBA.Len(.Description) > 0 Then Call CreateXmlElement(doc, "FullPath", .FullPath, , nd)
            If .IsBroken Then
                MsgBox .Name & " has a broken reference to: " & .Name, vbCritical
                Call CreateXmlElement(doc, "isBroken", .IsBroken, , nd)
            End If
        End With
    Next i
End Sub

Private Function isSelectedAddins() As Boolean  '// did user change installed Addins?
    Dim i As Long
    Dim n As Long
    
    For i = 1 To Excel.AddIns.Count
        If AddIns(i).Installed Then n = n + i
    Next i
    
    Debug.Print "Select Addins to Export Code"
    Application.Dialogs(xlDialogAddinManager).Show  '// .Dialogs(321).Show
    
    For i = 1 To Excel.AddIns.Count '// check to see if Addins were selected/deselected
        If AddIns(i).Installed Then n = n - i
    Next i
    isSelectedAddins = (n <> 0)
    
End Function

Private Function isVBEPermissionsOn() As Boolean
    On Error Resume Next
        If Not Application.VBE.VBProjects.Count > 0 Then
            Debug.Print "Enable Trust Access to the VBE Project model"
            Application.CommandBars.ExecuteMso "MacroSecurity"  '// turn off macroSecurity
        '// Application.CommandBars.FindControl(ID:=3627).Execute  '//same thing
        Else
            Debug.Print "... already Trust Access to the VBE Project model"
        End If
    isVBEPermissionsOn = IsNumeric(Application.VBE.VBProjects.Count)
End Function

Private Function isVBEPermissionsOff() As Boolean
    Debug.Print "Disable Trust Access to the VBA Project model for safety"
    Application.CommandBars.ExecuteMso "MacroSecurity"
    On Error Resume Next
    Debug.Assert IsNumeric(Application.VBE.VBProjects.Count)
    isVBEPermissionsOff = (Err.Number = 1004)
End Function
