Attribute VB_Name = "DCRsltModule"
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

Function DCRsltIsSam(A As DCRslt) As Boolean
With A
If .ADif.Count > 0 Then Exit Function
If .BDif.Count > 0 Then Exit Function
If .AExcess.Count > 0 Then Exit Function
If .BExcess.Count > 0 Then Exit Function
End With
DCRsltIsSam = True
End Function

Function DCRsltLy(A As DCRslt, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
With A
Dim A1() As S1S2: A1 = DCRsltS1S2Ay_Of_AExcess(.AExcess)
Dim A2() As S1S2: A2 = DCRsltS1S2Ay_Of_BExcess(.BExcess)
Dim A3() As S1S2: A3 = DCRsltS1S2Ay_Of_Dif(.ADif, .BDif)
Dim A4() As S1S2: A4 = DCRsltS1S2Ay_Of_Sam(.Sam)
End With
Dim O() As S1S2
S1S2_Push O, NewS1S2(Nm1, Nm2)
O = S1S2_Add(O, A1)
O = S1S2_Add(O, A2)
O = S1S2_Add(O, A3)
O = S1S2_Add(O, A4)
DCRsltLy = S1S2Ay_FmtLy(O)
End Function

Function DCRsltS1S2Ay_Of_AExcess(AExcess As Dictionary) As S1S2()
If Dix(AExcess).IsEmp Then Exit Function
Dim O() As S1S2, K
For Each K In AExcess.Keys
    S1S2_Push O, NewS1S2(K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & AExcess(K), "")
Next
DCRsltS1S2Ay_Of_AExcess = O
End Function

Function DCRsltS1S2Ay_Of_BExcess(BExcess As Dictionary) As S1S2()
If Dix(BExcess).IsEmp Then Exit Function
Dim O() As S1S2, K
For Each K In BExcess.Keys
    S1S2_Push O, NewS1S2("", K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & BExcess(K))
Next
DCRsltS1S2Ay_Of_BExcess = O
End Function

Function DCRsltS1S2Ay_Of_Dif(ADif As Dictionary, BDif As Dictionary) As S1S2()
Dim A As Dix: Set A = Dix(ADif)
Dim B As Dix: Set B = Dix(BDif)
If A.N <> B.N Then Stop
If A.IsEmp Then Exit Function
Dim O() As S1S2, K, S1$, S2$
For Each K In ADif
    S1 = K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & ADif(K)
    S2 = K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & BDif(K)
    S1S2_Push O, NewS1S2(S1, S2)
Next
DCRsltS1S2Ay_Of_Dif = O
End Function

Function DCRsltS1S2Ay_Of_Sam(ASam As Dictionary) As S1S2()
If Dix(ASam).IsEmp Then Exit Function
Dim O() As S1S2, K
For Each K In ASam.Keys
    S1S2_Push O, NewS1S2("*Same", K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & ASam(K))
Next
DCRsltS1S2Ay_Of_Sam = O
End Function
