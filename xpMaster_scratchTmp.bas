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
    Dim sh As Excel.Shape
    Dim co As Excel.ChartObject
    Dim sr As Excel.ShapeRange
    Dim cm As Excel.Comment
    Dim i As Long
    
    Set ws = ActiveSheet
''    Stop
    For i = 1 To ws.Shapes.Count
        Set sh = ws.Shapes(i)
        Debug.Print "----"; i
        Debug.Print , "[ID]:"; sh.ID
        Debug.Print , "[Name]: "; sh.Name
        Debug.Print , "[TopLeftCell]: "; sh.TopLeftCell.Address
        Debug.Print , "[BottomRightCell]: "; sh.BottomRightCell.Address
        Debug.Print , "[L,T,H,W:]: "; sh.Left; sh.Top; sh.Height; sh.Width
        Debug.Print , "[ZorderPos]:"; sh.ZOrderPosition
        Debug.Print , "[hasChart]:"; sh.HasChart
        Debug.Print , "[AltText]: '"; sh.AlternativeText; "'"
        Debug.Print "----"
        
        Select Case sh.Type
        Case MsoShapeType.msoChart
            Debug.Assert sh.HasChart
            Debug.Print , "[Chart.Name]:"; sh.Chart.Name
            sh.Chart.Export FileName:=sh.Chart.Name & ".png", FilterName:="png"
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
