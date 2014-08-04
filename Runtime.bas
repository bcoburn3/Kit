Attribute VB_Name = "Module1"
Global GlobalEnv As New Scripting.Dictionary
Global CurrentRes As Variant
Global LambdaCounter As Integer

Public Function FunCall(instance, ParamArray args())
    Set LR = GlobalEnv.item(instance.GetVal)
    FuncName = LR.item(1)
    Dim func As InventorVBAMember, TestFunc As InventorVBAMember
    Dim project As InventorVBAProject, module As InventorVBAComponent
    For Each project In ThisApplication.VBAProjects
        For Each module In project.InventorVBAComponents
            For Each TestFunc In module.InventorVBAMembers
                If TestFunc.Name = FuncName Then
                    Set func = TestFunc
                    GoTo exitloop
                End If
            Next
        Next
    Next
    MsgBox ("function not found " & FuncName)
    Exit Function
exitloop:
    Dim ArgsCollection As New Collection
    ArgsCollection.Add instance
    For Each arg In args
        ArgsCollection.Add arg
    Next
    func.Arguments.item(1).Value = ArgsCollection
    Dim result As Variant
    Call func.Execute
    'MsgBox (result)
    Set FunCall = CurrentRes
End Function

Public Function Lookup(item, LocalEnv)
    If LocalEnv.Exists(item) Then
        Set Lookup = LocalEnv.item(item)
    Else
        Set Lookup = GlobalEnv.item(item)
    End If
End Function

Public Function CreateLVal(val, kind)
    Dim res As New Lval
    If InStr("number, string, func, symb", kind) <> 0 Then
        Call res.SetVal(val, kind)
    Else
        Call res.SetRef(val, kind)
    End If
    Set Create = res
End Function

Public Function Cons(val1 As Lval, val2 As Lval)
    Dim res As New List
    Set res.Car = val1
    Set res.Cdr = val2
    Set Cons = res
End Function
