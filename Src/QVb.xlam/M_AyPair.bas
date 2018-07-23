Attribute VB_Name = "M_AyPair"
Option Explicit

Function AyPair_Dic(A1, A2) As Dictionary
Dim N1&, N2&
N1 = Sz(A1)
N2 = Sz(A2)
If N1 <> N2 Then Stop
Dim O As New Dictionary
Dim J&
If AyIsEmp(A1) Then GoTo X
For J = 0 To N1 - 1
    O.Add A1(J), A2(J)
Next
X:
Set AyPair_Dic = O
End Function
