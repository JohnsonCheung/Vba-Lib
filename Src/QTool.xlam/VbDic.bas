Attribute VB_Name = "VbDic"
Option Explicit
Function DicAB_SamDic(A As Dictionary, B As Dictionary) As Dictionary
Dim O As New Dictionary
If A.Count = 0 Or B.Count = 0 Then GoTo X
Dim K
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) = B(K) Then
            O.Add K, A(K)
        End If
    End If
Next
X: Set DicAB_SamDic = O
End Function
Function DicAB_SamKeyDifVal_DicPair(A As Dictionary, B As Dictionary) As Variant()
Dim K, A1 As New Dictionary, B1 As New Dictionary
For Each K In A.Keys
    If B.Exists(K) Then
        If A(K) <> B(K) Then
            A1.Add K, A(K)
            B1.Add K, B(K)
        End If
    End If
Next
DicAB_SamKeyDifVal_DicPair = Array(A1, B1)
End Function

Function DicAdd(A As Dictionary, B As Dictionary) As Dictionary
Dim O  As New Dictionary, I
For Each I In A.Keys
    O.Add I, A(I)
Next
For Each I In B.Keys
    O.Add I, B(I)
Next
Set DicAdd = O
End Function
Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, A(K)
Next
Set DicClone = O
End Function
Function DicCmp(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As DCRslt
Dim O As DCRslt
Set O.AExcess = DicMinus(A, B)
Set O.BExcess = DicMinus(B, A)
Set O.Sam = DicAB_SamDic(A, B)
Dim DicAB(): DicAB = DicAB_SamKeyDifVal_DicPair(A, B)
    Set O.ADif = DicAB(0)
    Set O.BDif = DicAB(1)
O.Nm1 = Nm1
O.Nm2 = Nm2
DicCmp = O
End Function
Function DicHasAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicHasAllKeyIsNm = True
End Function
Function DicHasAllValIsStr(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsStr(A(K)) Then Exit Function
Next
DicHasAllValIsStr = True
End Function
Function DicIsEq(A As Dictionary, B As Dictionary) As Boolean
Dim K(): K = A.Keys
If Sz(K) <> Sz(B.Keys) Then Exit Function
Dim KK, J%
For Each KK In K
    J = J + 1
    If KK = "*Dcl" Then
        If Len(A(KK)) <> Len(B(KK)) - 3 Then Stop
    Else
        If Len(A(KK)) <> Len(B(KK)) - 6 Then Stop
    End If
Next
DicIsEq = True
Stop
End Function
Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
If A.Count = 0 Then Set DicMinus = New Dictionary: Exit Function
If B.Count = 0 Then Set DicMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DicMinus = O
End Function
Function DicS1S2Itr(A As Dictionary) As Collection
Dim O As New Collection, K
For Each K In A.Keys
    O.Add S1S2(K, A(K))
Next
Set DicS1S2Itr = O
End Function
Function DicS1S2Ay(A As Dictionary) As S1S2()
Dim O() As S1S2, K
For Each K In A.Keys
    PushObj O, S1S2(K, A(K))
Next
DicS1S2Ay = O
End Function
Function DicSrt(A As Dictionary) As Dictionary
Dim Ky(): Ky = A.Keys
If Sz(Ky) = 0 Then Set DicSrt = New Dictionary: Exit Function
Dim Ky1(): Ky1 = AySrt(Ky)
Dim O As New Dictionary
Dim K
For Each K In Ky1
    O.Add K, A(K)
Next
Set DicSrt = O
End Function
Function DicWb(A As Dictionary, Optional Vis As Boolean) As Workbook
'Assume each dic keys is name and each value is lines
'Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Ass DicHasAllKeyIsNm(A)
Ass DicHasAllValIsStr(A)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook: Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        Set Ws = O.Sheets("Sheet1")
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = LinesSqV(A(K))
Next
X: Set Ws = O
If Vis Then O.Application.Visible = True
End Function
Function DicAy_Mge(A() As Dictionary) As Dictionary
'Assume there is no duplicated key in each of the dic in A()
Dim O As New Dictionary
If Sz(A) > 0 Then
    Dim I
    For Each I In A
        DicPush O, CvDic(I)
    Next
End If
Set DicAy_Mge = O
End Function
Sub DicPush(O As Dictionary, M As Dictionary)
'Assume there is no duplicated key
If M.Count = 0 Then Exit Sub
Dim K
For Each K In M.Keys
    O.Add K, M(K)
Next
End Sub
Function DicIsEmp(A As Dictionary) As Boolean
DicIsEmp = A.Count = 0
End Function
Function DicMap(A As Dictionary, ValMapFun$) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, Run(ValMapFun, A(K))
Next
Set DicMap = O
End Function
Function DicHasStrKy(A As Dictionary) As Boolean
DicHasStrKy = ItrPredAllTrue(A.Keys, "IsStr")
End Function
Function DicHasStrKy1(A As Dictionary) As Boolean
Dim I
For Each I In A.Keys
    If Not IsStr(I) Then Exit Function
Next
DicHasStrKy1 = True
End Function
Function DicAddKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set DicAddKeyPfx = O
End Function
Sub DicTyBrw(A As Dictionary)
DicBrw DicTy(A)
End Sub
Function DicTy(A As Dictionary) As Dictionary
Set DicTy = DicMap(A, "TyNm")
End Function
Sub DicWsBrw(A As Dictionary)
WsVis DicWs(A)
End Sub
Sub DicBrw(A As Dictionary)
Brw DicLy(A)
End Sub
Function DicLy(A As Dictionary) As String()
DicLy = S1S2AyFmt(DicS1S2Ay(A))
End Function
Function DicWs(A As Dictionary) As Worksheet
Set DicWs = S1S2AyWs(DicS1S2Ay(A))
End Function
Function DicStrKy(A As Dictionary) As String()
DicStrKy = AySy(A.Keys)
End Function
Function DicMaxValSz%(A As Dictionary)
'MthDic is DicOf_MthNm_zz_MthLinesAy
'MaxMthCnt is max-of-#-of-method per MthNm
Dim O%, K
For Each K In A.Keys
    O = Max(O, Sz(A(K)))
Next
DicMaxValSz = O
End Function
Function DicAyAdd(A() As Dictionary) As Dictionary
Dim O As New Dictionary, D
For Each D In A
    PushDic O, CvDic(D)
Next
Set DicAyAdd = O
End Function
Function DicA_RmvMth(A As Dictionary, MthNm$) As Dictionary
Dim O As New Dictionary
Dim K
For Each K In A.Keys
    If MthANm_MthNm(K) <> MthNm Then
        O.Add K, A(K)
    End If
Next
Set DicA_RmvMth = O
End Function
Private Sub ZZ_DicHasStrKy3()
TimFun "ZZ_DicHasStrKy ZZ_DicHasStrKy1"
End Sub
Private Sub ZZ_DicHasStrKy()
ZZ_DicHasStrKy__X "DicHasStrKy"
End Sub
Private Sub ZZ_DicHasStrKy1()
ZZ_DicHasStrKy__X "DicHasStrKy1"
End Sub
Private Sub ZZ_DicHasStrKy2()
Dim A As New Dictionary, Exp As Boolean, Act As Boolean
Dim J&
For J = 1 To 10000
    A.Add CStr(J), J
Next
Act = DicHasStrKy(A)
Exp = True
Ass Act = Exp

A.Add 10001, "X"
Act = DicHasStrKy(A)
Exp = False
Ass Act = Exp

End Sub
Private Sub ZZ_DicHasStrKy__X(X$)
Dim A As New Dictionary, Exp As Boolean, Act As Boolean
Dim J&
For J = 1 To 10000
    A.Add CStr(J), J
Next
Act = Run(X, A)
Exp = True
Ass Act = Exp

A.Add 10001, "X"
Act = Run(X, A)
Exp = False
Ass Act = Exp

End Sub
Private Sub ZZ_DicMaxValSz()
Dim D As Dictionary, M%
Set D = PjMthDic(CurPj)
M = DicMaxValSz(D)
Stop
End Sub
