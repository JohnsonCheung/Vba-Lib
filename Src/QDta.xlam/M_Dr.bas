Attribute VB_Name = "M_Dr"
Option Explicit
Function DrExpLinesCol(Dr, LinesColIx%) As Variant()
Dim B$()
    B = SplitCrLf(CStr(Dr(LinesColIx)))
Dim O()
    Dim IDr
        IDr = Dr
    Dim I
    For Each I In B
        IDr(LinesColIx) = I
        Push O, IDr
    Next
DrExpLinesCol = O
End Function

Property Get DrBySsl(Ssl$, TyAy() As eSimTy) As Variant()
Stop
End Property

Property Get DryBy_Ay_and_Const(Ay, Constant) As Variant()
If AyIsEmp(Ay) Then Exit Property
Dim O(), I
For Each I In Ay
   Push O, Array(I, Constant)
Next
DryBy_Ay_and_Const = O
End Property
Property Get DrLin$(Dr, Wdt%())
Dim UDr%
   UDr = UB(Dr)
Dim O$()
   Dim U1%: U1 = UB(Wdt)
   ReDim O(U1)
   Dim W, V
   Dim J%, V1$
   J = 0
   For Each W In Wdt
       If UDr >= J Then V = Dr(J) Else V = ""
       V1 = AlignL(V, W)
       O(J) = V1
       J = J + 1
   Next
DrLin = Quote(Join(O, " | "), "| * |")
End Property


Property Get DryBy_Const_and_Ay(Constant, Ay) As Variant()
If AyIsEmp(Ay) Then Exit Property
Dim O(), I
For Each I In Ay
   Push O, Array(Constant, I)
Next
DryBy_Const_and_Ay = O
End Property

Private Sub ZZ_InitByVblLy()
Dim VblLy$()
Dim Exp$()
Push VblLy, "|lskdf|sdlf|lsdkf"
Push VblLy, "|lsdf|"
Push VblLy, "|lskdfj|sdlfk|sdlkfj|sdklf|skldf|"
Push VblLy, "|sdf"
Dim Act As Drx
Set Act = InitByVblLy(VblLy)
Act.Brw
End Sub

Friend Property Get InitByS1S2s(S1S2s As S1S2s) As Drx
Dim Dry()
Dim J%, Ay() As S1S2
Ay = S1S2s.Ay
For J = 0 To S1S2s.U
   With Ay(J)
       Push Dry, Array(.S1, .S2)
   End With
Next
Set InitByS1S2s = Init(Dry)
End Property
Friend Property Get Init(Dry) As Drx
A = Dry
Set Init = Me
End Property
Friend Property Get InitByDic(A As Dictionary, Optional InclDicValTy As Boolean) As Drx
Stop
Set InitByDic = Me
End Property





Function DicDry(A As Dictionary, Optional InclDicValTy As Boolean) As Variant()
Dim O(), I
If A.Count = 0 Then Exit Function
Dim K(): K = A.Keys
If Not AyIsEmp(K) Then
   If InclDicValTy Then
       For Each I In K
           Push O, Array(I, A(I), TypeName(A(I)))
       Next
   Else
       For Each I In K
           Push O, Array(I, A(I))
       Next
   End If
End If
DicDry = O
End Function

Property Get AddConstCol(ConstVal) As Variant()
If IsEmp Then Exit Property
Dim O()
   Dim Dr, J&, NCol%
   NCol = Me.NCol
   ReDim O(URow)
   For Each Dr In A
       ReDim Preserve Dr(NCol)
       Dr(NCol) = ConstVal
       O(J) = Dr
       J = J + 1
   Next
AddConstCol = O
End Property

Sub Brw(Optional MaxColWdt& = 100, Optional BrkColIx% = -1)
AyBrw Ly(MaxColWdt, BrkColIx)
End Sub

Property Get Col(Optional ColIx% = 0) As Variant()
If IsEmp Then Exit Property
Dim O(), Dr
For Each Dr In A
   Push O, Dr(ColIx)
Next
Col = O
End Property

Property Get ColSet(Col_Ix%) As Dictionary
Dim O As New Dictionary
If Not IsEmp Then
    Dim Dr
    For Each Dr In A
        SetPush O, Dr(Col_Ix)
    Next
End If
Set ColSet = O
End Property

Sub Dmp()
AyDmp Ly
End Sub

Private Property Get DrIx_IsBrk(Dr, DrIx&, BrkColIx%) As Boolean
If AyIsEmp(Dr) Then Exit Property
If DrIx = 0 Then Exit Property
If DrIx = UB(Dr) Then Exit Property
If Dr(DrIx)(BrkColIx) = Dr(DrIx - 1)(BrkColIx) Then Exit Property
DrIx_IsBrk = True
End Property

Property Get CvCellToStr(ShwZer As Boolean) As Variant()
Dim O(), Dr
For Each Dr In A
   Push O, AyCellSy(Dr, ShwZer)
Next
CvCellToStr = O
End Property

Property Get IntCol(ColIx%) As Integer()
IntCol = AyIntAy(Col(ColIx))
End Property

Property Get IsEq(B()) As Boolean
If NRow <> Sz(B) Then Exit Property
If NRow = 0 Then IsEq = True: Exit Property
Dim J&, Dr
For Each Dr In A
   If Not AyIsEq(Dr, B(J)) Then Exit Property
   J = J + 1
Next
IsEq = True
End Property

Property Get KeyGpAy(K_Ix%, Gp_Ix%) As Variant()
If IsEmp Then Exit Property
Dim J%, O, K, GpAy(), O_Ix&, Gp, Dr, K_Ay()
For Each Dr In A
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
KeyGpAy = O
End Property

Property Get Ly(Optional MaxColWdt& = 100, Optional BrkColIx% = -1, Optional ShwZer As Boolean) As String()
If IsEmp Then Exit Property
Dim A1()
    A1 = CvCellToStr(ShwZer)
Dim Hdr$
    Dim W%(): W = Drx(A1).WdtAy(MaxColWdt)
    If AyIsEmp(W) Then Exit Property
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
            IsBrk = DrIx_IsBrk(Dr, DrIx, BrkColIx)
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
Ly = O
End Property

Private Property Get DryLy_InsBrkLin(DryLy$(), ColIx%) As String()
If Sz(DryLy) = 2 Then DryLy_InsBrkLin = DryLy: Exit Property
Dim Hdr$: Hdr = DryLy(0)
Dim Fm&, L%
   Dim N%: N = ColIx + 1
   Dim P1&, P2&
   P1 = InstrN(Hdr, "|", N)
   P2 = InStr(P1 + 1, Hdr, "|")
   Fm = P1 + 1
   L = P2 - P1 - 1
Dim O$()
   Push O, DryLy(0)
   Dim LasV$: LasV = Mid(DryLy(1), Fm, L)
   Dim J&
   Dim V$
   For J = 1 To UB(DryLy) - 1
       V = Mid(DryLy(J), Fm, L)
       If LasV <> V Then
           Push O, Hdr
           LasV = V
       End If
       Push O, DryLy(J)
   Next
   Push O, AyLasEle(DryLy)
DryLy_InsBrkLin = O
End Property

Property Get Mge(MgeIx%, Sep$) As Variant()
Dim O(), J%
Dim Ix%
For J = 0 To URow
   Ix = Dry_MgeIx(O, A(J), MgeIx)
   If Ix = -1 Then
       Push O, A(J)
   Else
       O(Ix)(MgeIx) = O(Ix)(MgeIx) & Sep & A(J)(MgeIx)
   End If
Next
Mge = O
End Property

Private Property Get Dry_MgeIx&(Dry(), Dr, MgeIx%)
Dim O&, D, J%
For O = 0 To UB(Dry)
   D = Dry(O)
   For J = 0 To UB(Dr)
       If J <> MgeIx Then
           If Dr(J) <> D(J) Then GoTo Nxt
       End If
   Next
   Dry_MgeIx = O
   Exit Property
Nxt:
Next
Dry_MgeIx = -1
End Property
Property Get NRow&()
NRow = Sz(A)
End Property

Property Get NCol%()
Dim Dr, O%, M%
For Each Dr In A
   M = Sz(Dr)
   If M > O Then O = M
Next
NCol = O
End Property

Property Get ReOrd(PartialIxAy&()) As Variant()
If IsEmp Then Exit Property
Dim Dr, O()
For Each Dr In A
   Push O, AyReOrd(Dr, PartialIxAy)
Next
ReOrd = O
End Property

Function Rg(At As Range) As Range
Set Rg = SqRg(Sq, At)
End Function

Property Get RmvColByIxAy(IxAy%()) As Variant()
If IsEmp Then Exit Property
Dim O(), Dr
For Each Dr In A
   Push O, AyWhExclIxAy(Dr, IxAy)
Next
RmvColByIxAy = O
End Property

Property Get RowCnt&(ColIx&, EqVal)
If IsEmp Then Exit Property
Dim J&, O&, Dr
For Each Dr In A
   If Dr(ColIx) = EqVal Then O = O + 1
Next
RowCnt = O
End Property

Property Get Sel(ColIxAy&(), Optional CrtEmpColIfReqFldNotFound As Boolean) As Variant()
Dim O(), Dr
If IsEmp Then Exit Property
For Each Dr In A
   Push O, AyWhIxAy(Dr, ColIxAy, CrtEmpColIfReqFldNotFound)
Next
Sel = O
End Property

Property Get SelDis(ColIx%) As Variant()
If IsEmp Then Exit Property
Dim Dr, O()
For Each Dr In A
   PushNoDup O, Dr(ColIx)
Next
SelDis = O
End Property

Property Get SelDisIntCol(ColIx%) As Integer()
SelDisIntCol = AyIntAy(SelDis(ColIx))
End Property

Property Get Sq(Optional NCol0% = 0) As Variant()
If IsEmp Then Exit Property
Dim NRow&, NCol&
   If NCol0 <= 0 Then NCol = Me.NCol
   NRow = Me.NRow
Dim O()
   ReDim O(1 To NRow, 1 To NCol)
Dim C%, R&, Dr
   R = 0
   For Each Dr In A
       R = R + 1
       For C = 0 To Min(UB(Dr), NCol - 1)
           O(R, C + 1) = Dr(C)
       Next
   Next
Sq = O
End Property

Property Get Srt(ColIx%, Optional IsDes As Boolean) As Variant()
Dim Col: Col = Me.Col(ColIx)
Dim Ix&(): Ix = AySrtInToIxAy(Col, IsDes)
Dim J%, O()
For J = 0 To UB(Ix)
   Push O, A(Ix(J))
Next
Srt = O
End Property

Property Get StrCol(Dry, Optional ColIx% = 0) As String()
StrCol = AySy(Col(ColIx))
End Property

Property Get WdtAy(Dry, Optional MaxColWdt& = 100) As Integer()
Const CSub$ = "Drx.WdtAy"
If IsEmp Then Exit Property
Dim O%()
   Dim Dr, UDr%, U%, V, L%, J%
   U = -1
   For Each Dr In Dry
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
WdtAy = O
End Property

Property Get Wh(ColIx%, EqVal) As Variant()
Dim O()
Dim J&
For J = 0 To URow
   If A(J)(ColIx) = EqVal Then Push O, A(J)
Next
Wh = O
End Property

Property Get Ws(Dry, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis:=True)
Rg WsA1(O)
Set Ws = O
End Property

Friend Property Get InitByDotNy(DotNy$()) As Drx
If AyIsEmp(DotNy) Then Set InitByDotNy = Init(Array()): Exit Property
Dim O(), I
For Each I In DotNy
   With Brk1(I, ".")
       Push O, ApSy(.S1, .S2)
   End With
Next
Set InitByDotNy = Init(O)
End Property

