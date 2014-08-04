VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public LType As String
Public Lval As Variant

Public Function GetVal()
    GetVal = Lval
End Function

Public Function GetRef()
    Set GetRef = Lval
End Function

Public Function GetType()
    GetType = LType
End Function

Public Function SetVal(val, Optional kind = False)
    If kind <> False Then
        LType = kind
    End If
    Lval = val
End Function

Public Function SetRef(val, Optional kind = False)
    If kind <> False Then
        LType = kind
    End If
    Set Lval = val
End Function





