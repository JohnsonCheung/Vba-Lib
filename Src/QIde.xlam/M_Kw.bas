Attribute VB_Name = "M_Kw"
Option Explicit
Function KwIsFunTy(S) As Boolean
KwIsFunTy = AyHas(SyOf_FunTy, S)
End Function
Function KwIsMdy(Mdy) As Boolean
KwIsMdy = AyHas(Array("Private", "Public", "Friend", ""), Mdy)
End Function
Function KwIsMthTy(S) As Boolean
KwIsMthTy = AyHas(S, SyOf_MthTy)
End Function
