Attribute VB_Name = "A_Xls"
Option Explicit
Property Get P3LCFV(Lx%, Cno%, Fld$, Val) As P3LCFV
Dim O As New P3LCFV
With O
    .Lx = Lx
    .Cno = Cno
    .Fld = Fld
    .Val = Val
End With
Set P3LCFV = O
End Property
