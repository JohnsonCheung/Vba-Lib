Attribute VB_Name = "IdeMthSubZ"
Option Explicit
Function MdSubZLines$(A As CodeModule)
Dim Ny$()
Ny = AySrt(MdMthNy(A, WhMth(Nm:=WhNm("^Z_"))))
Dim O$()
Push O, ""
Push O, "Sub Z()"
PushAy O, Ny
Push O, "End Sub"
MdSubZLines = Join(O, vbCrLf)
End Function

