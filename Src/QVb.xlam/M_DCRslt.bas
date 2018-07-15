Attribute VB_Name = "M_DCRslt"
Option Explicit
Private Type DicPair
    A As Dictionary
    B As Dictionary
End Type

Property Get DCRslt_Ly(A As DCRslt, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As String()
Dim O() As S1S2
PushObj O, S1S2(Nm1, Nm2)
PushObjAy O, ZAExcess(A)
PushObjAy O, ZBExcess(A)
PushObjAy O, ZDif(A)
PushObjAy O, ZSam(A)
DCRslt_Ly = S1S2Ay_FmtLy(O)
End Property

Property Get DicCmp(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As DCRslt
Dim O As New DCRslt
Set O.AExcess = DicMinus(A, B)
Set O.BExcess = DicMinus(B, A)
Set O.Sam = Intersect(A, B)
With ZSamKeyDifVal_DicPair(A, B)
    Set O.ADif = .A
    Set O.BDif = .B
End With
O.Nm1 = Nm1
O.Nm2 = Nm2
Set DicCmp = O
End Property


Sub DCRslt_Brw(A As DCRslt, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
AyBrw DCRslt_Ly(A, Nm1, Nm2)
End Sub

Sub DicCmpBrw(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd")
DCRslt_Brw DicCmp(A, B, Nm1, Nm2)
End Sub

Private Property Get ZAExcess(A As DCRslt) As S1S2()
If A.AExcess.Count = 0 Then Exit Property
Dim O() As S1S2, K
For Each K In A.AExcess.Keys
    PushObj O, S1S2(K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & A.AExcess(K), "")
Next
ZAExcess = O
End Property

Private Property Get ZBExcess(A As DCRslt) As S1S2()
If A.BExcess.Count = 0 Then Exit Property
Dim O() As S1S2, K
For Each K In A.BExcess.Keys
    PushObj O, S1S2("", K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & A.BExcess(K))
Next
ZBExcess = O
End Property

Private Property Get ZDif(A As DCRslt) As S1S2()
With A
    If .ADif.Count = 0 And .BDif.Count = 0 Then Exit Property
    Dim O() As S1S2, K, S1$, S2$
    For Each K In .ADif
        S1 = K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & .ADif(K)
        S2 = K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & .BDif(K)
        PushObj O, S1S2(S1, S2)
    Next
End With
ZDif = O
End Property

Private Property Get ZIsSam(A As DCRslt) As Boolean
If A.ADif.Count > 0 Then Exit Property
If A.BDif.Count > 0 Then Exit Property
If A.AExcess.Count > 0 Then Exit Property
If A.BExcess.Count > 0 Then Exit Property
ZIsSam = True
End Property

Private Property Get ZSam(A As DCRslt) As S1S2()
If A.Sam.Count = 0 Then Exit Property
Dim O() As S1S2, K
For Each K In A.Sam.Keys
    PushObj O, S1S2("*Same", K & vbCrLf & StrDup(Len(K), "-") & vbCrLf & A.Sam(K))
Next
ZSam = O
End Property

Private Property Get ZSamKeyDifVal_DicPair(A As Dictionary, B As Dictionary) As DicPair
Dim K, A1 As New Dictionary, B1 As New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            A1.Add K, A(K)
            B1.Add K, B(K)
        End If
    End If
Next
With ZSamKeyDifVal_DicPair
    Set .A = A1
    Set .B = B1
End With
End Property

Private Sub ZZ_DCRslt_Brw()
Stop
End Sub
