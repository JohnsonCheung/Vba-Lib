Attribute VB_Name = "M_FmTo"
Option Explicit

Function FmTo_Cnt&(A As FmTo)
If FmTo_IsVdt(A) Then Exit Function
FmTo_Cnt = A.Toix - A.Fmix + 1
End Function

Function FmTo_HasU(A As FmTo, U&) As Boolean
If U < 0 Then Stop
If FmTo_IsVdt(A) Then Exit Function
If A.Fmix > U Then Exit Function
If A.Toix < U Then Exit Function
FmTo_HasU = True
End Function

Function FmTo_IsVdt(A As FmTo) As Boolean
FmTo_IsVdt = True
If A.Fmix < 0 Then Exit Function
If A.Toix < 0 Then Exit Function
If A.Fmix > A.Toix Then Exit Function
FmTo_IsVdt = False
End Function

Function FmTo_LnoCnt(A As FmTo) As LnoCnt
Dim Lno&, Cnt&
   Cnt = A.Toix - A.Fmix + 1
   If Cnt < 0 Then Cnt = 0
   Lno = A.Fmix + 1
With FmTo_LnoCnt
   .Cnt = Cnt
   .Lno = Lno
End With
End Function


