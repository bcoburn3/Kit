Public Function KitCreate(val As Variant, kind As String)
    Dim res As New Collection
    Call res.add(val)
    Call res.add(kind)
    Set KitCreate = res
End Function

Public Function KitLookup(key As String, env As KitEnv)
    If env.exists(key) Then
        Set KitLookup = env.lookup(key)
    Else
        Set KitLookup = KitLookup(key, env.Parent)
    End If
End Function

Public Function KitSet(key As String, val As KitVal, env As KitEnv)
    If env.exists(key) Then
        Set KitSet = env.setval(key, val)
    Else
        Set KitSet = KitSet(key, val, env.Parent)
    End If
End Function

Public Function KitCreateList(ParamArray args())
    Dim res As New KitList
    For i = UBound(args) To LBound(args) Step -1
        Call res.Push(args(i))
    Next
    Set KitCreateList = KitCreate(res, "list")
End Function

Public Function KitTrue(val As KitVal)
    If val.kind = "nil" Then
        KitTrue = False
    Else
        KitTrue = True
    End If
End Function

Public Function KitIf(test As KitVal, ifval As KitVal, elseval As KitVal)
    If KitTrue(test) = True Then
        Set KitIf = ifval
    Else
        Set KitIf = elseval
    End If
End Function

Public Function KitCopyEnv(env As KitEnv)
    Dim NewEnv As New KitEnv
    If Not env.Parent Is Nothing Then
        Set NewEnv.Parent = KitCopyEnv(env.Parent)
    Else
        Set NewEnv.Parent = Nothing
    End If
    For Each key In env.Dict.Keys
        Call NewEnv.setval(key, env.lookup(key))
    Next
    Set KitCopyEnv = NewEnv
End Function

Public Function KitFuncall(func As KitVal, ParamArray args())
    Dim LR As Collection
    Set LR = func.val
    Dim env As New KitEnv
    Set env.Parent = LR.item(2)
    names = Split(LR.item(3), ", ")
    For i = 1 To UBound(names)
        Call env.setval(KitCreate(names(i), "symb"), args(i))
    Next
    Dim FuncObj As InventorVBAMember
    If LR.count = 4 Then
        Set FuncObj = LR.item(4)
    Else
        For Each project In ThisApplication.VBAProjects
            For Each module In project.InventorVBAComponents
                For Each TestFunc In module.InventorVBAMembers
                    If TestFunc.Name = LR.item(1) Then
                        Set FuncObj = TestFunc
                        GoTo exitloop
                    End If
                Next
            Next
        Next
exitloop:
    End If
    FuncObj.Arguments.item(1).value = env
    Dim result As New KitVal
    FuncObj.Execute (result)
    Set KitFuncall = result
End Function
