Attribute VB_Name = "testJSstaticClass"
Option Explicit

Sub testJS()
    Dim arr, o, v, x
    
    Debug.Print JS.epoch(0)
    Set arr = JS.parse("[3,5,7,11]")
    Debug.Print JS.stringify(arr)
    Set o = JS.parse("{""k3"":""v3"",""k2"": {""kk1"":""vv1"",""kk2"":""vv2""}}")
    Debug.Print JS.stringify(o)
    Debug.Print JS.arrayPush(arr, o)    '// push o pointer to arr
    Debug.Print JS.stringify(arr)
    Debug.Print JS(arr, "[4].k2.kk1")   '// hierarchical referencing
    Debug.Print TypeName(JS(arr, "[4].k2.xxx"))   '// 'Empty', not found
    JS.arrayPop arr, v  '// object pointer
    Debug.Print JS.stringify(v)
    Debug.Print JS.arrayPush(arr, o)
    Debug.Print JS.addItem(o, "k9", "v9")
    Debug.Print JS.stringify(o)
    Debug.Print JS.stringify(arr)
    Debug.Print JS.addItem(o, "nullkey", Null)
    Debug.Print JS.stringify(o, "", "    ")
    
End Sub

