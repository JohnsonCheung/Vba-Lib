Attribute VB_Name = "M_MthBrk"
Option Explicit
Function MthBrk_Str$(A As MthBrk)
Dim O$()
PushNonEmp O, A.Mdy
PushNonEmp O, A.Ty
PushNonEmp O, A.MthNm
MthBrk_Str = JnSpc(O)
End Function
