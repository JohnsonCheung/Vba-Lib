Attribute VB_Name = "DtaDry"
Option Explicit
Function DryColSeqCntDic(A, ColIx) As Dictionary
Set DryColSeqCntDic = AySeqCntDic(DryCol(A, ColIx))
End Function
Function DryAddValIdCntCol(A, ColIx) As Variant() ' Add 2 col at end (Id and Cnt) according to col(ColIx)
Dim O(), NCol%, Dr, R&, D As Dictionary, UCol%, IdCnt&()
O = A
UCol = DryNCol(O) + 1   ' The UCol after add
Set D = DryColSeqCntDic(A, ColIx)
For Each Dr In A
    ReDim Preserve Dr(UCol)
    If Not D.Exists(Dr(ColIx)) Then Stop
    IdCnt = D(ColIx)
    Dr(UCol - 1) = IdCnt(0)
    Dr(UCol) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
DryAddValIdCntCol = O
End Function

Function DryAddCol(A(), C) As Variant()
Dim UCol%, R&, Dr, O()
O = AyReSz(O, A)
UCol = DryNCol(A)
For Each Dr In AyNz(A)
    ReDim Preserve Dr(UCol)
    Dr(UCol) = C
    O(R) = Dr
    R = R + 1
Next
DryAddCol = O
End Function
Function DryAddCC(A(), C1, C2) As Variant()
DryAddCC = DryAddCol(DryAddCol(A, C1), C2)
End Function
Function DryAddValIdCol(A(), ValCol) As Variant()
Dim NCol%, Dic As Dictionary, O(), Dr, IdCnt, R&
NCol = DryNCol(A)
Set Dic = AyDistIdCntDic(DryCol(A, ValCol))
O = AyReSz(O, A)
For Each Dr In AyNz(A)
    ReDim Preserve Dr(NCol + 1)
    IdCnt = Dic(Dr(ValCol))
    Dr(NCol) = IdCnt(0)
    Dr(NCol + 1) = IdCnt(1)
    O(R) = Dr
    R = R + 1
Next
DryAddValIdCol = O
End Function
Function DryInsCol(A, C, Optional Ix&) As Variant()
Dim Dr
For Each Dr In A
    PushI DryInsCol, AyIns(Dr, C, At:=Ix)
Next
End Function
Function DryWhColEq(A, C%, V) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C) = V Then PushI DryWhColEq, Dr
Next
End Function
Function DryWhColGt(A, C%, V) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    If Dr(C) > V Then PushI DryWhColGt, Dr
Next
End Function
Function DryCntDic(A, KeyColIx%) As Dictionary
Dim O As New Dictionary
Dim J%, Dr, K
For J = 0 To UB(A)
    Dr = A(J)
    K = Dr(KeyColIx)
    If O.Exists(K) Then
        O(K) = O(K) + 1
    Else
        O.Add K, 1
    End If
Next
Set DryCntDic = O
End Function
Function DryNCol&(A())
If Sz(A) = 0 Then Exit Function
Dim O&, Dr
For Each Dr In A
    O = Max(O, Sz(Dr))
Next
DryNCol = O
End Function
Function DrySq(A()) As Variant()
If Sz(A) = 0 Then Exit Function
Dim NCol&, NRow&
    NCol = DryNCol(A)
    NRow = Sz(A)
Dim O()
ReDim O(1 To NRow, 1 To NCol)
Dim C&, R&, Dr
    For R = 1 To NRow
        Dr = A(R - 1)
        For C = 1 To Min(Sz(Dr), NCol)
            O(R, C) = Dr(C - 1)
        Next
    Next
DrySq = O
End Function
Function DryStrCol(A(), ColIx) As String()
DryStrCol = DryColInto(A, ColIx, EmpSy)
End Function
Function DryColInto(A, ColIx, OInto)
Dim O, J&, Dr, U&
O = AyReSz(OInto, A)
For Each Dr In AyNz(A)
    If UB(Dr) >= ColIx Then
        O(J) = Dr(ColIx)
    End If
    J = J + 1
Next
DryColInto = O
End Function
Function DryCol(A, ColIx) As Variant()
DryCol = DryColInto(A, ColIx, Array())
End Function
Function DryWhDup(A, Optional ColIx% = 0) As Variant()
Dim Dup, Dr, O()
Dup = AyWhDup(DryCol(A, ColIx))
For Each Dr In A
    If AyHas(Dup, Dr(ColIx)) Then Push O, Dr
Next
DryWhDup = O
End Function
Function DryWdt(A()) As Integer()
Dim J%
For J = 0 To DryNCol(A) - 1
    Push DryWdt, AyWdt(DryCol(A, J))
Next
End Function

Sub DrsFmtssBrw(A As Drs)
Brw DrsFmtss(A)
End Sub

Sub DrsFmtssDmp(A As Drs)
D DrsFmtss(A)
End Sub

Function DrsFmtss(A As Drs) As String()
DrsFmtss = DryFmtss(CvAy(ItmAddAy(A.Fny, A.Dry)))
End Function

Function DryFmtss(A()) As String()
Dim W%(), Dr, O$()
W = DryWdt(A)
For Each Dr In AyNz(A)
    PushI O, DrFmtss(Dr, W)
Next
DryFmtss = O
End Function
Function DryFmtssWrp(A(), Optional WrpWdt% = 40) As String() _
'WrpWdt is for wrp-col.  If maxWdt of an ele of wrp-col > WrpWdt, use the maxWdt
Dim W%(), Dr, A1(), M$()
W = WrpDryWdt(A, WrpWdt)
For Each Dr In AyNz(A)
    M = DrFmtssWrp(Dr, W)
    PushIAy DryFmtssWrp, M
Next
End Function

Function DryFmtssCell(A()) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    Push DryFmtssCell, DrFmtssCell(Dr)
Next
End Function
Function DryGpDic(A, K%, G%) As Dictionary
Dim Dr, U&, O As New Dictionary, KK, GG, Ay()
U = UB(A): If U = -1 Then Exit Function
For Each Dr In A
    KK = Dr(K)
    GG = Dr(G)
    If O.Exists(KK) Then
        Ay = O(KK)
        Push Ay, GG
        O(KK) = Ay
    Else
        O.Add KK, Array(GG)
    End If
Next
Set DryGpDic = O
End Function
Function DryRmvCol(A, ColIx&) As Variant()
Dim X
For Each X In AyNz(A)
    PushI DryRmvCol, AyRmvEleAt(X, ColIx)
Next
End Function
Function DryGpFlat(A, K%, G%) As Variant()
DryGpFlat = Aydic_to_KeyCntMulItmColDry(DryGpDic(A, K, G))
End Function
Function DryWs(A()) As Worksheet
Set DryWs = SqWs(DrySq(A))
End Function
Function DryInsCCC(A, C1, C2, C3) As Variant()
Dim Dr, O(), C3Ay()
If Sz(A) = 0 Then Exit Function
C3Ay = Array(C1, C2, C3)
For Each Dr In A
    Push O, AyInsAy(Dr, C3Ay)
Next
DryInsCCC = O
End Function
Function DryInsC4(A, C1, C2, C3, C4) As Variant()
Dim Dr, O(), C4Ay()
If Sz(A) = 0 Then Exit Function
C4Ay = Array(C1, C2, C3, C4)
For Each Dr In A
    Push O, AyInsAy(Dr, C4Ay)
Next
DryInsC4 = O
End Function
Function DryInsCC(A, C1, C2) As Variant()
Dim Dr, O(), CCAy()
If Sz(A) = 0 Then Exit Function
CCAy = Array(C1, C2)
For Each Dr In A
    Push O, AyInsAy(Dr, CCAy)
Next
DryInsCC = O
End Function
