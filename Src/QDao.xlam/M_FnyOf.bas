Attribute VB_Name = "M_FnyOf"
Option Explicit
Function FnyOf_FldInf() As String()
FnyOf_FldInf = SplitSpc("Fld Pk Ty Sz Dft Req Des")
End Function
Function FnyOf_TblFInf() As String()
Dim O$()
Push O, "Tbl"
Push O, "SeqNo"
PushAy O, FnyOf_FldInf
FnyOf_TblFInf = O
End Function
