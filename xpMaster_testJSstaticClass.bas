Attribute VB_Name = "testJSstaticClass"
Option Explicit


'Function to mediate clicks on worksheet hyperlinks
Function LinkController(Url)
    Debug.Print "URL: ", Url
    Set LinkController = Selection 'have to return something...
    If Url Like "*goog*" Then
        ThisWorkbook.FollowHyperlink Url
    Else
        MsgBox "Can't open this link"
    End If
End Function

Sub testJS()
    Dim arr, o, v, x
    Dim js
    
    Debug.Print js.epoch(0)
    Set arr = js.parse("[3,5,7,11]")
    Debug.Print js.stringify(arr)
    Set o = js.parse("{""k3"":""v3"",""k2"": {""kk1"":""vv1"",""kk2"":""vv2""}}")
    Debug.Print js.stringify(o)
    Debug.Print js.arrayPush(arr, o)    '// push o pointer to arr
    Debug.Print js.stringify(arr)
    Debug.Print js(arr, "[4].k2.kk1")   '// hierarchical referencing
    Debug.Print TypeName(js(arr, "[4].k2.xxx"))   '// 'Empty', not found
    js.arrayPop arr, v  '// object pointer
    Debug.Print js.stringify(v)
    Debug.Print js.arrayPush(arr, o)
    Debug.Print js.addItem(o, "k9", "v9")
    Debug.Print js.stringify(o)
    Debug.Print js.stringify(arr)
    Debug.Print js.addItem(o, "nullkey", Null)
    Debug.Print js.stringify(o, "", "    ")
    
End Sub

