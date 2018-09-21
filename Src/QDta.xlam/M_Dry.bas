Attribute VB_Name = "M_Dry"
Option Explicit

Function DryAddConstCol(Dry(), ConstVal) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim N%
   N = Sz(Dry(0))
Dim O()
   Dim Dr, J&
   ReDim O(UB(Dry))
   For Each Dr In Dry
       ReDim Preserve Dr(N)
       Dr(N) = ConstVal
       O(J) = Dr
       J = J + 1
   Next
DryAddConstCol = O
End Function

Function DryCol(Dry, Optional ColIx% = 0) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim O(), Dr
For Each Dr In Dry
   Push O, Dr(ColIx)
Next
DryCol = O
End Function

Function DryColColl(A, ColIx%) As Collection
Dim O As New Collection
If Not Sz(A) = 0 Then
    Dim Dr
    For Each Dr In A
        O.Add Dr(ColIx)
    Next
End If
Set DryColColl = O
End Function

Function DryGpAy(A, Kix%, Gix%) As Variant()
If Sz(A) = 0 Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In A
    K = Dr(Kix)
    Gp = Dr(Gix)
    O_Ix = AyIx(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryGpAy = O
End Function

Function DryIntCol(A, ColIx%) As Integer()
DryIntCol = DryCol_Into(A, ColIx, EmpIntAy)
End Function

Function DryIsBrkAtDrIx(Dry, DrIx&, BrkColIx%) As Boolean
If AyIsEmp(Dry) Then Exit Function
If DrIx = 0 Then Exit Function
If DrIx = UB(Dry) Then Exit Function
If Dry(DrIx)(BrkColIx) = Dry(DrIx - 1)(BrkColIx) Then Exit Function
DryIsBrkAtDrIx = True
End Function

Function DryIsEq(A(), B()) As Boolean
Dim N&: N = Sz(A)
If N <> Sz(B) Then Exit Function
If N = 0 Then DryIsEq = True: Exit Function
Dim J&, Dr
For Each Dr In A
   If Not AyIsEq(Dr, B(J)) Then Exit Function
   J = J + 1
Next
DryIsEq = True
End Function

Function DryKeyGpAy(Dry(), K_Ix%, Gp_Ix%) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In Dry
    K = Dr(K_Ix)
    Gp = Dr(Gp_Ix)
    O_Ix = AyIx(K_Ay, K)
    If O_Ix = -1 Then
        Push K_Ay, K
        Push O, Array(K, Array(Gp))
    Else
        Push O(O_Ix)(1), Gp
    End If
Next
DryKeyGpAy = O
End Function
Sub ZZ_DryLy()
AyDmp DryLy(SampleDry1)
End Sub
Function DryLy(A, Optional MaxColWdt& = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean) As String()
If AyIsEmp(A) Then Exit Function
Dim A1()
    A1 = DryStrCellDry(A, ShwZer)
Dim Hdr$
    Dim W%(): W = DryWdtAy(A1, MaxColWdt)
    If AyIsEmp(W) Then Exit Function
    Dim HdrAy$()
    ReDim HdrAy(UB(W))
    Dim J%
    For J = 0 To UB(W)
        HdrAy(J) = String(W(J), "-")
    Next
    Hdr = Quote(Join(HdrAy, "-|-"), "|-*-|")

Dim O$()
    Dim Dr, DrIx&, IsBrk As Boolean
    Push O, Hdr
    If BrkColIx >= 0 Then
        For Each Dr In A1
            IsBrk = DryIsBrkAtDrIx(A, DrIx, BrkColIx)
            If IsBrk Then Push O, Hdr
            Push O, DrLin(Dr, W)
            DrIx = DrIx + 1
        Next
    Else
        For Each Dr In A1
            Push O, DrLin(Dr, W)
        Next
    End If
    Push O, Hdr
DryLy = O
End Function

Function DryMge(Dry, MgeIx%, Sep$) As Variant()
Dim O(), J%
Dim Ix%
For J = 0 To UB(Dry)
   Ix = DryMgeIx(O, Dry(J), MgeIx)
   If Ix = -1 Then
       Push O, Dry(J)
   Else
       O(Ix)(MgeIx) = O(Ix)(MgeIx) & Sep & Dry(J)(MgeIx)
   End If
Next
DryMge = O
End Function

Function DryMgeIx&(Dry, Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   DryMgeIx = O
   Exit Function
Nxt:
Next
DryMgeIx = -1
End Function

Function DryNCol%(Dry)
Dim Dr, O%, M%
For Each Dr In Dry
   M = Sz(Dr)
   If M > O Then O = M
Next
DryNCol = O
End Function

Function DryReOrd(Dry, PartialIxAy&()) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim Dr, O()
For Each Dr In Dry
   Push O, AyReOrd(Dr, PartialIxAy)
Next
DryReOrd = O
End Function

Function DryRg(A, At As Range) As Range
Set DryRg = SqRg(DrySq(A), At)
End Function

Function DryRmvColByIxAy(Dry, IxAy%()) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim O(), Dr
For Each Dr In Dry
   Push O, AyWhExclIxAy(Dr, IxAy)
Next
DryRmvColByIxAy = O
End Function

Function DryRowCnt&(Dry, ColIx&, EqVal)
If AyIsEmp(Dry) Then Exit Function
Dim J&, O&, Dr
For Each Dr In Dry
   If Dr(ColIx) = EqVal Then O = O + 1
Next
DryRowCnt = O
End Function

Function DrySel(A(), CIxAy&(), Optional CrtEmpCol_IfReqCol_NotFound As Boolean) As Variant()
Dim O(), Dr
If Sz(A) = 0 Then Exit Function
For Each Dr In A
   Push O, AyWhIxAy(Dr, CIxAy, CrtEmpCol_IfReqCol_NotFound)
Next
DrySel = O
End Function

Function DrySelDis(A(), ColIx%) As Variant()
If Sz(A) = 0 Then Exit Function
Dim Dr, O()
For Each Dr In A
   PushNoDup O, Dr(ColIx)
Next
DrySelDis = O
End Function

Function DrySelDisIntCol(Dry(), ColIx%) As Integer()
DrySelDisIntCol = AyIntAy(DrySelDis(Dry, ColIx))
End Function

Function DrySq(Dry, Optional NColOpt% = 0) As Variant()
If AyIsEmp(Dry) Then Exit Function
Dim NRow&, NCol&
   If NColOpt <= 0 Then NCol = DryNCol(Dry)
   NRow = Sz(Dry)
Dim O()
   ReDim O(1 To NRow, 1 To NCol)
Dim C%, R&, Dr
   R = 0
   For Each Dr In Dry
       R = R + 1
       For C = 0 To UB(Dr)
           O(R, C + 1) = Dr(C)
       Next
   Next
DrySq = O
End Function

Function DrySrt(Dry, ColIx%, Optional IsDes As Boolean) As Variant()
Dim Col: Col = DryCol(Dry, ColIx)
Dim Ix&(): Ix = AySrtInToIxAy(Col, IsDes)
Dim J%, O()
For J = 0 To UB(Ix)
   Push O, Dry(Ix(J))
Next
DrySrt = O
End Function

Function DryStrCellDry(A, ShwZer As Boolean) As Variant()
Dim O(), Dr
If AyIsEmp(A) Then Exit Function
For Each Dr In A
   Push O, AyCellSy(Dr, ShwZer)
Next
DryStrCellDry = O
End Function

Function DryStrCol(A, Optional ColIx% = 0) As String()
DryStrCol = DryCol_Into(A, ColIx, EmpSy)
End Function

Function DryStrDry(A, ShwZer As Boolean) As Variant()
Dim O(), Dr
For Each Dr In A
   Push O, AyCellSy(Dr, ShwZer)
Next
DryStrDry = O
End Function

Function DryWdtAy(A, Optional MaxColWdt& = 100) As Integer()
Const CSub$ = "DryWdtAy"
If Sz(A) = 0 Then Exit Function
Dim O%()
   Dim Dr, UDr%, U%, V, L%, J%
   U = -1
   For Each Dr In A
       If Not IsSy(Dr) Then Er CSub, "This routine should call DryCvFmtEachCell first so that each cell is ValCellStr as a string.|Now some Dr in given-Dry is not a StrAy, but[" & TypeName(Dr) & "]"
       UDr = UB(Dr)
       If UDr > U Then ReDim Preserve O(UDr): U = UDr
       If AyIsEmp(Dr) Then GoTo Nxt
       For J = 0 To UDr
           V = Dr(J)
           L = Len(V)

           If L > O(J) Then O(J) = L
       Next
Nxt:
   Next
Dim M%
M = MaxColWdt
For J = 0 To UB(O)
   If O(J) > M Then O(J) = M
Next
DryWdtAy = O
End Function

Function DryWh(A, ColIx%, EqVal) As Variant()
Dim O()
Dim J&
For J = 0 To UB(A)
   If A(J)(ColIx) = EqVal Then Push O, A(J)
Next
DryWh = O
End Function

Function DryWs(Dry, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
DryRg Dry, WsA1(O)
Set DryWs = O
End Function

Sub DryBrw(Dry, Optional MaxColWdt& = 100, Optional BrkColIx% = -1)
AyBrw DryLy(Dry, MaxColWdt, BrkColIx)
End Sub

Sub DryDmp(Dry)
AyDmp DryLy(Dry)
End Sub

Private Function Dry_MgeIx&(Dry(), Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   Dry_MgeIx = O
   Exit Function
Nxt:
Next
Dry_MgeIx = -1
End Function
