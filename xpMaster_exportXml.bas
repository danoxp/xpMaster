Attribute VB_Name = "exportXml"
Option Explicit

Public Sub ExportXMLsheets()   '// exports both forms of XML for all non blank sheets in ActiveWorkbook\
    Dim wb As Excel.Workbook
    Dim ai As Excel.AddIn
    Dim cm As Office.COMAddIn
    Dim sh As Excel.Worksheet
    Dim folderName As String
    
    With CreateObject("scripting.filesystemobject")
        For Each wb In Excel.Application.Workbooks
            If Not wb.Saved Then MsgBox wb.Name & " not saved, skipped": Exit For
            folderName = wb.Path & "\" & Replace(wb.Name, ".xls", "_")
            For Each sh In wb.Worksheets
                If Not IsEmpty(sh.UsedRange) Then    '// if Not a blank sheet export xml
                    .CreateTextFile(folderName & sh.Name & "_excel.xml").Write sh.UsedRange.Value(xlRangeValueXMLSpreadsheet)
                    Debug.Print (folderName & sh.Name & "_excel.xml")
                    .CreateTextFile(folderName & sh.Name & "_MSpersist.xml").Write sh.UsedRange.Value(xlRangeValueMSPersistXML)
                    Debug.Print (folderName & sh.Name & "_MSpersist.xml")
                End If
            Next sh
        Next wb
        
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


