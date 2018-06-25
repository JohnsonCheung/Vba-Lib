Attribute VB_Name = "M_FmTo"
Option Explicit

Property Get FmTo_Cnt&(A As FmTo)
If FmTo_IsVdt(A) Then Exit Property
FmTo_Cnt = A.ToIx - A.FmIx + 1
End Property

Property Get FmTo_HasU(A As FmTo, U&) As Boolean
If U < 0 Then Stop
If FmTo_IsVdt(A) Then Exit Property
If A.FmIx > U Then Exit Property
If A.ToIx < U Then Exit Property
FmTo_HasU = True
End Property

Property Get FmTo_IsVdt(A As FmTo) As Boolean
FmTo_IsVdt = True
If A.FmIx < 0 Then Exit Property
If A.ToIx < 0 Then Exit Property
If A.FmIx > A.ToIx Then Exit Property
FmTo_IsVdt = False
End Property

Property Get FmTo_LnoCnt(A As FmTo) As LnoCnt
Dim Lno&, Cnt&
   Cnt = A.ToIx - A.FmIx + 1
   If Cnt < 0 Then Cnt = 0
   Lno = A.FmIx + 1
With FmTo_LnoCnt
   .Cnt = Cnt
   .Lno = Lno
End With
End Property
