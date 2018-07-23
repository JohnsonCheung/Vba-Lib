Attribute VB_Name = "M_SyOf"
Option Explicit
Private Const C_Ty$ = "Type"
Private Const C_Enm$ = "Enum"
Private Const C_Get$ = "Get"
Private Const C_Let$ = "Let"
Private Const C_Set$ = "Set"
Private Const C_Fun$ = "Function"
Private Const C_Sub$ = "Sub"
Private Const C_Prp$ = "Property"
Private Const C_PrpLet$ = "Property Let"
Private Const C_PrpGet$ = "Property Get"
Private Const C_PrpSet$ = "Property Set"
Private Const C_Pub$ = "Public"
Private Const C_Prv$ = "Private"
Private Const C_Frd = "Friend"
Function SyOf_FunTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Fun, C_Sub, C_Prp)
End If
SyOf_FunTy = Y
End Function
Function SyOf_Mdy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Pub, C_Prv, C_Frd)
End If
SyOf_Mdy = Y
End Function
Function SyOf_MthTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Fun, C_Sub, C_PrpGet, C_PrpLet, C_PrpSet)
End If
SyOf_MthTy = Y
End Function
Function SyOf_PrpTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Get, C_Set, C_Let)
End If
SyOf_PrpTy = Y
End Function
Function SyOf_SrcTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = SyOf_MthTy
   Push Y, C_Ty
   Push Y, C_Enm
End If
SyOf_SrcTy = Y
End Function
