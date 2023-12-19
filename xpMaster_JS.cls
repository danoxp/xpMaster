VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
'// Attribute VB_PredeclaredId = True   '// Static 'Singleton' or global instance
'// Attribute VB_Exposed = True         '// access from outside VBAProject using Refernce
Option Explicit

Private m As thisClass
Private Type thisClass
    doc As Object   '// HTMLDocument
    win As Object   '// HTMLWindow2
    
    addItem As Object
    newObject As Object
    deleteItem  As Object
    keysObject As Object
    decUUT As Object
    an As Object
    anwd As Object
    fn As Object
    fnnn As Object
    getProp As Object
    getType As Object
    copyDict As Object
    
    JSON As Object
    pars As Object
    strfy As Object
    jsnow As Object
    clipboardData As Object
''    stack As Object
End Type

Private Sub Class_Initialize()
    Debug.Print "XP.xla-JS.Class_Initialize"
    Set m.doc = CreateObject("htmlfile")
    With m.doc
        .Open "text/html"
        .Close
        With .parentWindow
            Set m.win = .self
            '// check for IE11 functionality ... before loading functions
            .execScript "function isIE11(e) {return !!window.msCrypto}"
            If Not .isIE11(0) Then setRegistryExcelforIE11
            '// check ForceShellExecute HKEY_LOCAL_MACHINE\SOFTWARE\Wow6432Node\Microsoft\Office\9.0\Common\Internet
            '// check .NET CLR 3.5.30729 also HKLM check
            .execScript .atob(jsB64())
            Set m.addItem = .addItem
            Set m.newObject = .newObject
            Set m.deleteItem = .deleteItem
            Set m.keysObject = .keysObject
            Set m.decUUT = .decUUT
            Set m.an = .an
            Set m.anwd = .anwd
            Set m.fn = .fn
            Set m.fnnn = .fnnn
            Set m.getProp = .getProp
            Set m.getType = .getType
            Set m.copyDict = .copyDict
            '//
            Set m.JSON = .JSON
            Set m.pars = .JSON.parse
            Set m.strfy = .JSON.stringify
            Set m.jsnow = .Date.now ': Dim now      '// must be '.Date.now' lowercase
            Set m.clipboardData = .clipboardData
        End With
    End With
End Sub

Private Sub Class_Terminate()
    Debug.Print "XP.xla-JS.Class_Terminate"
    Set m.doc = Nothing: Set m.win = Nothing
End Sub

Public Sub SetLet(A As Variant, B As Variant) '// set/let B = A
    If IsObject(A) Then Set B = A Else Let B = A
End Sub
Public Function arrayPush(arr As Variant, v As Variant) As Long  '// js array.push()
    On Error Resume Next
    If IsEmpty(arr) Then Set arr = JS.parse("[]")
    arrayPush = CallByName(arr, "push", VbMethod, v)
    
''    Select Case True
''    Case IsEmpty(arr)
''        Set arr = JS.parsejs("[]")
''        push = CallByName(arr, "push", VbMethod, v)
''    Case arr Is Nothing
''        Set arr = JS.parsejs("[]")
''        push = CallByName(arr, "push", VbMethod, v)
''    Case JS.typejs(arr) = "array"
''        push = CallByName(arr, "push", VbMethod, v)
''    Case Else
''        '// not JScriptTypeInfo = 'array'
''    End Select
    
End Function

Public Sub arrayPop(arr As Variant, ret As Variant) '// ret can be object, string, number, Emtpy, etc..
    SetLet CallByName(arr, "pop", VbMethod), ret
End Sub

Public Function htmlparse(jsonhtml As String)
Attribute htmlparse.VB_Description = "jsonhtml like '&quot;data&quot;:123&quot;' -> returns JScriptTypeInfo"
    On Error Resume Next
    m.doc.body.innerHTML = jsonhtml
    Set htmlparse = m.pars(m.doc.body.innerText)
    If Not Err.Number = 0 Then Set htmlparse = Nothing
End Function

Public Property Get confirmjs()         '// returns true|false
Attribute confirmjs.VB_Description = "returns true|false"
    Set confirmjs = m.win.confirm
End Property

Public Property Get clipboard() As String
    clipboard = m.clipboardData.getData("text") ': Dim getData
End Property
Public Property Let clipboard(s As String)
        If Len(s) > 0 Then m.clipboardData.setData "text", s
End Property

Public Property Get window()
    Set window = m.win
End Property

Public Property Get document()
    Set document = m.doc
End Property

Public Property Get epoch()
    Set epoch = m.jsnow
End Property

Public Property Get alert()
    Set alert = m.win.alert
End Property

Public Property Get DateObject()     '// set o = getDateObject(1581096379000) -> Date object
    Set DateObject = m.win.getDate   '// o.toUTCString(), o.getTimezoneOffset(), o.toISOString()
End Property

Public Property Get Datejs()    '// local time -> 'Sun Feb 27 2022 11:23:30 GMT-0500 (Eastern Standard Time)'
    Set Datejs = m.win.Date
End Property

Public Function DateParse()     '// DateParse("Fri, 7 Feb 2020 10:26:19 -0700") -> 1581096379000 (double/dec)
    Set DateParse = m.win.Date.parse
End Function

Public Property Get encodeURIComponent()
    Set encodeURIComponent = m.win.encodeURIComponent
End Property

Public Property Get decodeURIComponent()
    Set decodeURIComponent = m.win.decodeURIComponent
End Property

Public Property Get encodeURI()
    Set encodeURI = m.win.encodeURI
End Property

Public Property Get decodeURI()
    Set decodeURI = m.win.decodeURI
End Property

Public Property Get atob()
    Set atob = m.win.atob
End Property

Public Property Get btoa()
    Set btoa = m.win.btoa
End Property

'// non-native functions
Public Property Get addItem()
Attribute addItem.VB_Description = "decodes 'u003Du002F -> '= , '"
    Set addItem = m.addItem
End Property

Public Property Get newObject()
    Set newObject = m.newObject
End Property

Public Property Get deleteItem()
    Set deleteItem = m.deleteItem
End Property

Public Property Get keysObject()
    Set keysObject = m.keysObject
End Property

Public Property Get decodeUUTF()        '// '/u003D' -> '='
    Set decodeUUTF = m.decUUT
End Property

Public Property Get allNodesWithData()
    Set allNodesWithData = m.anwd
End Property

Public Property Get allNodes()
    Set allNodes = m.an
End Property

Public Property Get firstNodeNotNull()
    Set firstNodeNotNull = m.fnnn
End Property

Public Property Get deepitem()
    Set deepitem = m.fn
End Property

Public Property Get parse()
    Set parse = m.pars
End Property

Public Property Get stringify()
    Set stringify = m.strfy
End Property

Public Property Get itemjs()
Attribute itemjs.VB_Description = "Call like 'itemjs(object, key)' -> returns value"
Attribute itemjs.VB_UserMemId = 0
    '// Default Class member
    Set itemjs = m.getProp
End Property

Public Property Get typejs()
    Set typejs = m.getType
End Property

Private Function jsB64()
    jsB64 = _
    "ZnVuY3Rpb24gYWRkSXRlbShlLHQsbil7cmV0dXJuIGVbdF09bixPYmplY3Qua2V5cyhlKS5sZW5ndGh9ZnVuY3Rpb24gbmV3T2JqZWN0KGUsdCl7dmFyIG49e307cmV0dXJuIDE9PT1hcmd1bWVudHMubGVuZ3RoPzAhPT1lJiYobltlXT1udWxsKToyPT09YXJndW1lbnRzLmxlbmd0aC" & _
    "YmKG5bZV09dCksbn1mdW5jdGlvbiBkZWxldGVJdGVtKGUsdCl7cmV0dXJuIGRlbGV0ZSBlW3RdLE9iamVjdC5rZXlzKGUpLmxlbmd0aH1mdW5jdGlvbiBrZXlzT2JqZWN0KGUpe3JldHVybiBPYmplY3Qua2V5cyhlKX1mdW5jdGlvbiBnZXREYXRlKGUpe3JldHVybiBuZXcgRGF0ZShl" & _
    "KX1mdW5jdGlvbiBkZWNVVVQoZSl7cmV0dXJuIEpTT04ucGFyc2UoJyInK2UucmVwbGFjZSgvXHUiL2csJ1xcdSInKSsnIicpfWZ1bmN0aW9uIGFuKGUsdCxuKXt2YXIgcj1ufHxbXTtmb3IodmFyIGwgaW4gZSlsIT09dD9udWxsIT09ZVtsXSYmIm9iamVjdCI9PXR5cGVvZiBlW2xdJi" & _
    "ZPYmplY3Qua2V5cyhlW2xdKS5sZW5ndGg+MCYmYW4oZVtsXSx0LHIpOnIucHVzaChlW2xdKTtyZXR1cm4gcn1mdW5jdGlvbiBhbndkKGUsdCxuKXt2YXIgcj1ufHxbXTtmb3IodmFyIGwgaW4gZSlsIT09dD9udWxsIT09ZVtsXSYmIm9iamVjdCI9PXR5cGVvZiBlW2xdJiZPYmplY3Qu" & _
    "a2V5cyhlW2xdKS5sZW5ndGg+MCYmYW53ZChlW2xdLHQscik6KCJvYmplY3QiIT10eXBlb2YgZVtsXXx8bnVsbCE9PWVbbF0mJk9iamVjdC5rZXlzKGVbbF0pLmxlbmd0aD4wKSYmci5wdXNoKGVbbF0pO3JldHVybiByfWZ1bmN0aW9uIGZuKGUsdCxuKXtzd2l0Y2goITApe2Nhc2UgZS" & _
    "5oYXNPd25Qcm9wZXJ0eSh0KTpyZXR1cm4gZVt0XTtjYXNlIW4mJnQuc2VhcmNoKC9cLnxcWy8pPi0xOnZhciByPXQucmVwbGFjZSgvXlxbfFxdL2csIiIpLnJlcGxhY2UoL1xbL2csIi4iKS5zcGxpdCgiLiIpO2xldCBsPWU7Zm9yKGxldCBlPTA7ZTxyLmxlbmd0aDtlKyspe2lmKG51" & _
    "bGw9PT1sfHwhbC5oYXNPd25Qcm9wZXJ0eShyW2VdKSlyZXR1cm47bD1sW3JbZV1dfXJldHVybiBsO2RlZmF1bHQ6bGV0IHU7Zm9yKGxldCBuIGluIGUpaWYobnVsbCE9PWVbbl0mJiJvYmplY3QiPT10eXBlb2YgZVtuXSYmT2JqZWN0LmtleXMoZVtuXSkubGVuZ3RoPjAmJih1PWZuKG" & _
    "Vbbl0sdCwhMCksdm9pZCAwIT09dSkpcmV0dXJuIHV9fWZ1bmN0aW9uIGZpcnN0Tm9kZU5vdE51bGwoZSx0KXtpZihlLmhhc093blByb3BlcnR5KHQpJiZudWxsIT09ZVt0XSlyZXR1cm4gZVt0XTt2YXIgbjtmb3IodmFyIHIgaW4gZSlpZihudWxsIT09ZVtyXSYmIm9iamVjdCI9PXR5" & _
    "cGVvZiBlW3JdJiZPYmplY3Qua2V5cyhlW3JdKS5sZW5ndGg+MCYmdm9pZCAwIT09KG49Zmlyc3ROb2RlTm90TnVsbChlW3JdLHQpKSlyZXR1cm4gbn1mdW5jdGlvbiBmbm5uKGUsdCl7aWYoZS5oYXNPd25Qcm9wZXJ0eSh0KSYmbnVsbCE9PWVbdF0pcmV0dXJuIGVbdF07dmFyIG47Zm" & _
    "9yKHZhciByIGluIGUpaWYobnVsbCE9PWVbcl0mJiJvYmplY3QiPT10eXBlb2YgZVtyXSYmT2JqZWN0LmtleXMoZVtyXSkubGVuZ3RoPjAmJnZvaWQgMCE9PShuPWZubm4oZVtyXSx0KSkpcmV0dXJuIG59ZnVuY3Rpb24gZ2V0UHJvcChlLHQpe3N3aXRjaCghMCl7Y2FzZSBlLmhhc093" & _
    "blByb3BlcnR5KHQpOnJldHVybiBlW3RdO2Nhc2UtMT09PXQuc2VhcmNoKC9cLnxcWy8pOnJldHVybjtkZWZhdWx0OnZhciBuPXQucmVwbGFjZSgvXlxbfFxdL2csIiIpLnJlcGxhY2UoL1xbL2csIi4iKS5zcGxpdCgiLiIpO2xldCByPWU7Zm9yKGxldCBlPTA7ZTxuLmxlbmd0aDtlKy" & _
    "spe2lmKG51bGw9PT1yfHwhci5oYXNPd25Qcm9wZXJ0eShuW2VdKSlyZXR1cm47cj1yW25bZV1dfXJldHVybiByfX1mdW5jdGlvbiBnZXRUeXBlKGUpe3N3aXRjaCh0eXBlb2YgZSl7Y2FzZSJzdHJpbmciOmNhc2UibnVtYmVyIjpjYXNlImJvb2xlYW4iOmNhc2UibnVsbCI6cmV0dXJu" & _
    "IHR5cGVvZiBlO2Nhc2Uib2JqZWN0IjppZighZSlyZXR1cm4ibnVsbCI7aWYoIltvYmplY3QgQXJyYXldIj09PU9iamVjdC5wcm90b3R5cGUudG9TdHJpbmcuYXBwbHkoZSkpcmV0dXJuImFycmF5In1yZXR1cm4ib2JqZWN0In1mdW5jdGlvbiBjb3B5RGljdChlLHQpe2Zvcih2YXIgbi" & _
    "BpbiBlKXQuQWRkKG4sZVtuXSk7cmV0dXJuIHR9"
End Function

Private Sub setRegistryExcelforIE11()   '// set H
    Const HKCU = &H80000001
    Const REGKEY = "SOFTWARE\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION"
    Dim ret

    With GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")
        ret = .SetDWORDValue(HKCU, REGKEY, "Excel.exe", 0)
    End With
    If ret = 0 Then
        m.win.alert ("Registry setting enabling Excel.exe update -> reboot Excel")
        Stop: Application.Quit
    Else
        MsgBox "Error trying to set Excel.exe registry setting -> dano"
    End If

End Sub
