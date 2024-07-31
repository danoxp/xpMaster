Attribute VB_Name = "exportXml"
Option Explicit

Public Sub ExportXMLsheets()   '// exports both forms of XML for all non blank sheets in ActiveWorkbook\
    Dim Wb As Excel.Workbook
    Dim ai As Excel.AddIn
    Dim cm As Office.COMAddIn
    Dim Sh As Excel.Worksheet
    Dim folderName As String
    
    With CreateObject("scripting.filesystemobject")
        For Each Wb In Excel.Application.Workbooks
            If Not Wb.Saved Then MsgBox Wb.Name & " not saved, skipped": Exit For
            folderName = Wb.Path & "\" & Replace(Wb.Name, ".xls", "_")
            For Each Sh In Wb.Worksheets
                If Not IsEmpty(Sh.UsedRange) Then    '// if Not a blank sheet export xml
                    .CreateTextFile(folderName & Sh.Name & "_excel.xml").Write Sh.UsedRange.Value(xlRangeValueXMLSpreadsheet)
                    Debug.Print (folderName & Sh.Name & "_excel.xml")
                    .CreateTextFile(folderName & Sh.Name & "_MSpersist.xml").Write Sh.UsedRange.Value(xlRangeValueMSPersistXML)
                    Debug.Print (folderName & Sh.Name & "_MSpersist.xml")
                End If
            Next Sh
        Next Wb
        
        For Each ai In Excel.AddIns
            Debug.Print ai.FullName
            If ai.Installed Then
                Debug.Print Workbooks(ai.Name).Sheets.Count
                Debug.Print Workbooks(ai.Name).VBProject.VBComponents.Count
            End If
        Next ai
        
        For Each cm In Application.COMAddIns
            Debug.Print cm.Description
        Next cm
        
    End With
End Sub


