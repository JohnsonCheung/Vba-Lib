Attribute VB_Name = "M_FnyOf"
Option Explicit
Function FnyzDbInfzFld() As String()
FnyzDbInfzFld = SplitSpc("Fld Pk Ty Sz Dft Req Des")
End Function
Function FnyzDbInfzTbl() As String()
Dim O$()
Push O, "Tbl"
Push O, "SeqNo"
PushAy O, FnyzDbInfzFld
FnyzDbInfzTbl = O
End Function
Function FnyzDbInfzTblF() As String()
Dim O$()
FnyzDbInfzTblF = O
End Function

