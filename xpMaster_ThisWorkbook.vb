VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private WithEvents oEvents As clsEvents
Attribute oEvents.VB_VarHelpID = -1

'// oEvents, clsEvents - Begin
Public Sub TurnOnAppEvents()
    Set oEvents = New clsEvents
End Sub

Public Sub TurnOffAppEvents()
    Set oEvents = Nothing
End Sub

Private Sub oEvents_MySheetChange()
    Debug.Print "ThisWorkbook: oEvents_MySheetChange"
''    VBA.DoEvents  '// this didn't work???s
    Debug.Print "ThisWorkbook: oEvents_MySheetChange-END:"; oEvents.N
End Sub
'// oEvents, clsEvents - End

Public Sub testAccess(s1 As String, s2 As String)
    Debug.Print "testAccess", s1, s2
End Sub

Public Sub testAccessN(N As Long)
    Debug.Print "testAccessN:"; N
End Sub


Private Sub Workbook_Open()
    Debug.Print "event: ", ThisWorkbook.Name, "ThisWorkbook.Workbook_Open"
''    Application.CommandBars(1).Reset
''    Main.XpSearchOff        'Turn off XpSearch Addin on excel boot?
''    Main.installXpControl
    With Application
        .OnKey "{F1}", "Main.xpBuiltInMenusPopup"       'F1 All Excel Menus
        .OnKey "{F3}", "Main.xpF3"                      'F3 Next Window
        .OnKey "{F5}", "Main.xpF5"                      'F5 FollowHyperlink, SaveChartPngFile, XpSearch
        .OnKey "{F6}", "Main.xpF6"                      'F6 toggle AutoFilter FreezeTopRow
        .OnKey "{F7}", "Main.xpF7"                      'F7 usedrange
    End With
''    Application.OnTime now, "'" & ThisWorkbook.FullName & "'!Main.initEvents"
''    Main.initEvents
End Sub

Private Sub Workbook_AddinUninstall()
    Debug.Print "event: ", ThisWorkbook.Name, "ThisWorkbook.Workbook_AddinUninstall"
    MsgBox "AddinUninstall"
    Main.deleteXPcontrols
    With Application
        .OnKey "{F1}"
        .OnKey "{F3}"
        .OnKey "{F5}"
        .OnKey "{F6}"
        .OnKey "{F7}"
    End With
''    Main.killEvents
End Sub
