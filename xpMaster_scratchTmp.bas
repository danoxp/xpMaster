Attribute VB_Name = "scratchTmp"
Option Explicit

Sub makeHyperlinks()
    Dim i As Long
    
    With ActiveSheet
        For i = 1 To .UsedRange.Rows.Count
''            If Not IsEmpty(Cells(i, 2)) Then
            If Cells(i, 2).Hyperlinks.Count = 0 Then
                With Cells(i, 1).Hyperlinks.Add(Cells(i, 1), Cells(i, 2).Value)
                    .ScreenTip = "Profile | LinkedIn"
                End With
            Else
''                With Cells(i, 2).Hyperlinks.Add(Cells(i, 2), "https://www.google.com/search?q=site%3Alinkedin.com%2Fin+")
                With Cells(i, 3).Hyperlinks.Add(Cells(i, 3), "https://www.bing.com/search?q=site%3Alinkedin.com%2Fin%2F+")
                    .Address = .Address & Cells(i, 1).Value
''                    .TextToDisplay = "google"
                    .TextToDisplay = "bing"
                End With
            End If
        Next i
    End With
End Sub

Sub chartObjectTesting()
    Dim ws As Excel.Worksheet
    Dim Sh As Excel.Shape
    Dim co As Excel.ChartObject
    Dim sr As Excel.ShapeRange
    Dim cm As Excel.Comment
    Dim i As Long
    
    Set ws = ActiveSheet
''    Stop
    For i = 1 To ws.Shapes.Count
        Set Sh = ws.Shapes(i)
        Debug.Print "----"; i
        Debug.Print , "[ID]:"; Sh.ID
        Debug.Print , "[Name]: "; Sh.Name
        Debug.Print , "[TopLeftCell]: "; Sh.TopLeftCell.Address
        Debug.Print , "[BottomRightCell]: "; Sh.BottomRightCell.Address
        Debug.Print , "[L,T,H,W:]: "; Sh.Left; Sh.Top; Sh.Height; Sh.Width
        Debug.Print , "[ZorderPos]:"; Sh.ZOrderPosition
        Debug.Print , "[hasChart]:"; Sh.HasChart
        Debug.Print , "[AltText]: '"; Sh.AlternativeText; "'"
        Debug.Print "----"
        
        Select Case Sh.Type
        Case MsoShapeType.msoChart
            Debug.Assert Sh.HasChart
            Debug.Print , "[Chart.Name]:"; Sh.Chart.Name
            Sh.Chart.Export FileName:=Sh.Chart.Name & ".png", FilterName:="png"
            Stop
        Case MsoShapeType.msoComment
            Stop
        Case MsoShapeType.msoPicture
            Stop
        Case MsoShapeType.msoShapeTypeMixed
            Stop
        Case MsoShapeType.msoTextBox
            Stop
        Case Else
            Debug.Assert False
        End Select
        
    Next i
    Stop
End Sub
