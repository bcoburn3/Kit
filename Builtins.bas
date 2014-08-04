Attribute VB_Name = "Builtins"
Public Sub BuiltinInit()
    Dim NullEnv As New Scripting.Dictionary
    Dim PlusLR As New Collection
    PlusLR.Add "plus"
    PlusLR.Add NullEnv
    Set GlobalEnv.item("+LR") = PlusLR
    Set GlobalEnv.item("+") = Create("+LR", "func")
    
    Dim MinusLR As New Collection
    MinusLR.Add "minus"
    MinusLR.Add NullEnv
    Set GlobalEnv.item("-LR") = MinusLR
    Set GlobalEnv.item("-") = Create("-LR", "func")
    
    Dim TimesLR As New Collection
    TimesLR.Add "times"
    TimesLR.Add NullEnv
    Set GlobalEnv.item("*LR") = TimesLR
    Set GlobalEnv.item("*") = Create("*LR", "func")
    
    Dim DivideLR As New Collection
    DivideLR.Add "divide"
    DivideLR.Add NullEnv
    Set GlobalEnv.item("/LR") = DivideLR
    Set GlobalEnv.item("/") = Create("/LR", "func")
    
    Dim ConsLR As New Collection
    ConsLR.Add "Cons"
    ConsLR.Add NullEnv
    Set GlobalEnv.item("ConsLR") = ConsLR
    Set GlobalEnv.item("Cons") = Create("ConsLR", "func")
    
    Dim CarLR As New Collection
    CarLR.Add "Car"
    CarLR.Add NullEnv
    Set GlobalEnv.item("CarLR") = CarLR
    Set GlobalEnv.item("Car") = Create("CarLR", "func")
    
    Dim CdrLR As New Collection
    CdrLR.Add "Cdr"
    CdrLR.Add NullEnv
    Set GlobalEnv.item("CdrLR") = CdrLR
    Set GlobalEnv.item("Cdr") = Create("CdrLR", "func")
    
End Sub

Public Function Cons(args As Collection)
'function prelude
    instance = args.item(1).GetVal
    Dim LR As New Collection
    Set LR = GlobalEnv.item(instance)
    Dim Env As New Scripting.Dictionary
    Set Env = LR.item(2)
'actual function body
Dim res As New List
Set res.Car = args.item(2)
Set res.Cdr = args.item(3)
'local environment writeback
    Dim LR_Writeback As New Collection
    Call LR_Writeback.Add(LR.item(1))
    Call LR_Writeback.Add(Env)
    Set GlobalEnv.item(instance) = LR_Writeback
'return value
    Set Module1.current_res = res
End Function

 Public Function Car(args As Collection)
'function prelude
    instance = args.item(1).GetVal
    Dim LR As New Collection
    Set LR = GlobalEnv.item(instance)
    Dim Env As New Scripting.Dictionary
    Set Env = LR.item(2)
'actual function body
Set res = args.item(2).Car
'local environment writeback
    Dim LR_Writeback As New Collection
    Call LR_Writeback.Add(LR.item(1))
    Call LR_Writeback.Add(Env)
    Set GlobalEnv.item(instance) = LR_Writeback
'return value
    Set Module1.current_res = res
End Function

 Public Function Cdr(args As Collection)
'function prelude
    instance = args.item(1).GetVal
    Dim LR As New Collection
    Set LR = GlobalEnv.item(instance)
    Dim Env As New Scripting.Dictionary
    Set Env = LR.item(2)
'actual function body
Set res = args.item(2).Cdr
'local environment writeback
    Dim LR_Writeback As New Collection
    Call LR_Writeback.Add(LR.item(1))
    Call LR_Writeback.Add(Env)
    Set GlobalEnv.item(instance) = LR_Writeback
'return value
    Set Module1.current_res = res
End Function


Public Function plus(args As Collection)
'function prelude
    instance = args.item(1).GetVal
    Dim LR As New Collection
    Set LR = GlobalEnv.item(instance)
    Dim Env As New Scripting.Dictionary
    Set Env = LR.item(2)
'actual function body
    Dim res As New Lval
    If args.Count = 1 Then
        Set res = Create(0, "number")
    ElseIf args.Count = 2 Then
        Set res = Create(args.item(2).GetVal, "number")
    Else
        Set res = Create(args.item(2).GetVal, "number")
        For i = 3 To args.Count
            Call res.SetVal((res.GetVal + args.item(i).GetVal), "number")
        Next
    End If
'captured variable writeback
    Dim LR_Writeback As New Collection
    Call LR_Writeback.Add(LR.item(1)) 'add function ID, even though we don't need it
    Call LR_Writeback.Add(Env)
    Set GlobalEnv.item(instance) = LR_Writeback
'return value
    Set Module1.current_res = res
End Function

Public Function minus(args As Collection)
'function prelude
    instance = args.item(1).GetVal
    Dim LR As New Collection
    Set LR = GlobalEnv.item(instance)
    Dim Env As New Scripting.Dictionary
    Set Env = LR.item(2)
'actual function body
    Dim res As New Lval
    If args.Count = 1 Then
        Set res = Create(0, "number")
    ElseIf args.Count = 2 Then
        Set res = Create(0 - args.item(2).GetVal, "number")
    Else
        Set res = Create(args.item(2).GetVal, "number")
        For i = 3 To args.Count
            Call res.SetVal((res.GetVal - args.item(i).GetVal), "number")
        Next
    End If
'captured variable writeback
    Dim LR_Writeback As New Collection
    Call LR_Writeback.Add(LR.item(1)) 'add function ID, even though we don't need it
    Call LR_Writeback.Add(Env)
'return value
    Set Module1.current_res = res
End Function

Public Function times(args As Collection)
'function prelude
    instance = args.item(1).GetVal
    Dim LR As New Collection
    Set LR = GlobalEnv.item(instance)
    Dim Env As New Scripting.Dictionary
    Set Env = LR.item(2)
'actual function body
    Dim res As New Lval
    If args.Count = 1 Then
        Set res = Create(1, "number")
    ElseIf args.Count = 2 Then
        Set res = Create(args.item(2).GetVal, "number")
    Else
        Set res = Create(args.item(2).GetVal, "number")
        For i = 3 To args.Count
            Call res.SetVal((res.GetVal * args.item(i).GetVal), "number")
        Next
    End If
'captured variable writeback
    Dim LR_Writeback As New Collection
    Call LR_Writeback.Add(LR.item(1)) 'add function ID, even though we don't need it
    Call LR_Writeback.Add(Env)
    Set GlobalEnv.item(instance) = LR_Writeback
'return value
    Set Module1.current_res = res
End Function

Public Function divide(args As Collection)
'function prelude
    instance = args.item(1).GetVal
    Dim LR As New Collection
    Set LR = GlobalEnv.item(instance)
    Dim Env As New Scripting.Dictionary
    Set Env = LR.item(2)
'actual function body
    Dim res As New Lval
    If args.Count = 1 Then
        Set res = Create(1, "number")
    ElseIf args.Count = 2 Then
        Set res = Create((1 / args.item(2).GetVal), "number")
    Else
        Set res = Create(args.item(2).GetVal, "number")
        For i = 3 To args.Count
            Call res.SetVal((res.GetVal / args.item(i).GetVal), "number")
        Next
    End If
'captured variable writeback
    Dim LR_Writeback As New Collection
    Call LR_Writeback.Add(LR.item(1)) 'add function ID, even though we don't need it
    Call LR_Writeback.Add(Env)
    Set GlobalEnv.item(instance) = LR_Writeback
'return value
    Set Module1.current_res = res
End Function
