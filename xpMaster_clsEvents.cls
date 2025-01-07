VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents XL As Excel.Application
Attribute XL.VB_VarHelpID = -1

Public Event MySheetChange()

Public col As Collection
Public N As Long

Private Sub Class_Initialize()
    Debug.Print "XP.xla: Events.Class_Initialize"
    Set XL = Application
End Sub

Private Sub Class_Terminate()
    Debug.Print "XP.xla: Events.Class_Terminate"
    Set XL = Nothing
End Sub

Private Sub XL_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    Debug.Print "XP.xla: Events.XL_SheetChange"; N; "Sh: "; Sh.Name, "Target: "; Target.Address
    If N = 0 Then
''        DoEvents
        RaiseEvent MySheetChange
    '// RaiseEvent is no different than a 'Method Call'
    Application.OnTime Now(), "'ThisWorkbook.testAccess " & Chr$(34) & "Test1" & Chr$(34) & "," & Chr$(34) & "Test2" & Chr$(34) & "'"
    Application.OnTime Now(), "'ThisWorkbook.testAccessN " & N & "'"
        
    End If
''    VBA.DoEvents
    N = N + 1
    Debug.Print "SheetChangeEnd:"; N
End Sub
