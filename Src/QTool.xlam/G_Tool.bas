Attribute VB_Name = "G_Tool"
Option Explicit
Type Either
    IsLeft As Boolean
    Left As Variant
    Right As Variant
End Type
Type DCRslt
    Nm1 As String
    Nm2 As String
    AExcess As New Dictionary
    BExcess As New Dictionary
    ADif As New Dictionary
    BDif As New Dictionary
    Sam As New Dictionary
End Type
Type MdSrtRpt
    MdNy() As String
    RptDic As Dictionary ' K is Module Name, V is DicCmpRsltLy
End Type
Function VbeHasPj(A As Vbe, PjNm) As Boolean
VbeHasPj = ItrHasNm(A.VBProjects, PjNm)
End Function
Function ItrHasNm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then ItrHasNm = True: Exit Function
Next
End Function

Function TakBefOrNo$(S, Sep, Optional NoTrim As Boolean)
TakBefOrNo = Brk2(S, Sep, NoTrim).S1
End Function
Function TakAftOrNo$(S, Sep, Optional NoTrim As Boolean)
TakAftOrNo = Brk2(S, Sep, NoTrim).S2
End Function
Function TakAftMust$(A, Sep, Optional NoTrim As Boolean)
TakAftMust = Brk(A, Sep, NoTrim).S2
End Function
Function TakAft$(A, Sep, Optional NoTrim As Boolean)
TakAft = Brk2(A, Sep, NoTrim).S2
End Function
Function TakBef$(S, Sep$, Optional NoTrim As Boolean)
TakBef = Brk1(S, Sep, NoTrim).S1
End Function
Function TakBefMust$(S, Sep$, Optional NoTrim As Boolean)
TakBefMust = Brk(S, Sep, NoTrim).S1
End Function
Function Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk2 = Brk2__X(A, P, Sep, NoTrim)
End Function
Function Brk2__X(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk2__X = S1S2("", A)
    Else
        Set Brk2__X = S1S2("", Trim(A))
    End If
    Exit Function
End If
Set Brk2__X = BrkAt(A, P, Sep, NoTrim)
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
Function DryAddColByDic(A, KeyColIx%, Dic As Dictionary) As Variant()
Dim O(): O = A
Dim NCol%: NCol = DryNCol(O)
Dim Dr(), J&, V, K
For J = 0 To UB(A)
    Dr = A(J)
    ReDim Preserve Dr(NCol)
    K = Dr(KeyColIx)
    V = Dic(K)
    Dr(NCol) = V
    O(J) = Dr
Next
DryAddColByDic = O
End Function
Function AlignL$(A, W, Optional ErIFmnotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "AlignL"
If ErIFmnotEnoughWdt And DoNotCut Then
    Stop
    'Er CSub, "Both {ErIFmnotEnoughWdt} and {DontCut} cannot be True", ErIFmnotEnoughWdt, DoNotCut
End If
Dim S$: S = VarStr(A)
AlignL = StrAlignL(S, W, ErIFmnotEnoughWdt, DoNotCut)
End Function

Function AscIsDigit(A%) As Boolean
If A < 48 Then Exit Function
If A > 57 Then Exit Function
AscIsDigit = True
End Function

Function AscIsLCase(A%) As Boolean
If A < 97 Then Exit Function
If A > 122 Then Exit Function
AscIsLCase = True
End Function

Function AscIsUCase(A%) As Boolean
If A < 65 Then Exit Function
If A > 90 Then Exit Function
AscIsUCase = True
End Function
Function AyInto(A, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(A) > 0 Then
    Dim I
    For Each I In A
        Push O, I
    Next
End If
AyInto = O
End Function
Function AyAB_FmtLy(A, B) As String()
AyAB_FmtLy = S1S2Ay_FmtLy(AyAB_S1S2Ay(A, B))
End Function

Function AyAB_S1S2Ay(A, B) As S1S2()
Dim U&: U = UB(A)
If U <> UB(B) Then Stop
Dim O() As S1S2
ReDim O(U)
Dim J&
For J = 0 To U
    Set O(J) = S1S2(A(J), B(J))
Next
AyAB_S1S2Ay = O
End Function

Function AyAddAp(A, ParamArray Ap())
Dim O: O = A
Dim I
Dim Av(): Av = Ap
For Each I In Av
    PushAy O, I
Next
AyAddAp = O
End Function

Function AyAddPfx(A, Pfx) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & A(J)
Next
AyAddPfx = O
End Function

Function AyAddPfxSfx(A, Pfx, Sfx) As String()
Dim O$(), J&, U&
If Sz(A) = 0 Then Exit Function
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & A(J) & Sfx
Next
AyAddPfxSfx = O
End Function

Function AyAddSfx(A, Sfx) As String()
If Sz(A) = 0 Then Exit Function
Dim O$(), J&, U&
U = UB(A)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = A(J) & Sfx
Next
AyAddSfx = O
End Function

Function AyAlignL(Ay) As String()
Dim W%: W = AyWdt(Ay) + 1
If Sz(Ay) = 0 Then Exit Function
Dim O$(), I
For Each I In Ay
    Push O, AlignL(I, W)
Next
AyAlignL = O
End Function

Function AyCntDry(A) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), I, J&, Fnd As Boolean
For Each I In A
    Fnd = False
    For J = 0 To UB(O)
        If O(J)(0) = I Then
            O(J)(1) = O(J)(1) + 1
            Fnd = True
            Exit For
        End If
    Next
    If Not Fnd Then
        Push O, Array(I, 1)
    End If
Next
AyCntDry = O
End Function

Function AyItr(A) As Collection
Dim O As New Collection, I
If Sz(A) = 0 Then Set AyItr = O: Exit Function
For Each I In A
    O.Add I
Next
Set AyItr = O
End Function

Function AyDblQuote(A) As String()
Const C$ = """"
AyDblQuote = AyAddPfxSfx(A, C, C)
End Function

Function AyFstNEle(A, N&)
Dim O: O = A
ReDim Preserve O(N - 1)
AyFstNEle = O
End Function

Function AyHas(A, M) As Boolean
Dim I: If Sz(A) = 0 Then Exit Function
For Each I In A
    If I = M Then AyHas = True: Exit Function
Next
End Function

Function AyIns(A, Optional M, Optional At&)
Dim N&: N = Sz(A)
If 0 > At Or At > N Then
    Stop
End If
Dim O
    O = A
    ReDim Preserve O(N)
    Dim J&
    For J = N To At + 1 Step -1
        Asg O(J - 1), O(J)
    Next
    O(At) = M
AyIns = O
End Function

Function AyIsAllEleEq(A) As Boolean
If Sz(A) = 0 Then AyIsAllEleEq = True: Exit Function
Dim J&
For J = 1 To UB(A)
    If A(0) <> A(J) Then Exit Function
Next
AyIsAllEleEq = True
End Function

Function AyIsEq(A1, A2) As Boolean
Dim U&: U = UB(A1): If U <> UB(A2) Then Exit Function
Dim J&
For J = 0 To U
   If A1(J) <> A2(J) Then Exit Function
Next
AyIsEq = True
End Function

Function AyIx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then AyIx = J: Exit Function
Next
AyIx = -1
End Function

Function AyLasEle(Ay)
AyLasEle = Ay(UB(Ay))
End Function

Function AyMap(A, MapFunNm$)
AyMap = AyMapInto(A, MapFunNm, EmpAy)
End Function

Function AyMapInto(A, MapFunNm$, OIntoAy)
Dim O: O = OIntoAy: Erase O
Dim I
If Sz(A) > 0 Then
    For Each I In A
        Push O, Run(MapFunNm, I)
    Next
End If
AyMapInto = O
End Function

Function AyMapPXInto(A, MapPXFunNm$, P, OIntoAy)
'MapPXFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
' or "Mth"
Dim O: O = OIntoAy: Erase O
Dim X
If Sz(A) > 0 Then
    For Each X In A
        Push O, Run(MapPXFunNm, P, X)
    Next
End If
AyMapPXInto = O
End Function
Function AyMapXPInto(A, MapXPFunNm$, P, OIntoAy)
'MapXPFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
' or "Mth"
Dim O: O = OIntoAy: Erase O
Dim X
If Sz(A) > 0 Then
    For Each X In A
        Push O, Run(MapXPFunNm, X, P)
    Next
End If
AyMapXPInto = O
End Function

Function AyMapPXSy(A, MapPXFunNm$, Prm) As String()
AyMapPXSy = AyMapPXInto(A, MapPXFunNm, Prm, EmpSy)
End Function
Function AyMapXPSy(A, MapXPFunNm$, Prm) As String()
AyMapXPSy = AyMapXPInto(A, MapXPFunNm, Prm, EmpSy)
End Function
Function AyMapXP(A, MapXPFunNm$, Prm) As Variant()
AyMapXP = AyMapXPInto(A, MapXPFunNm, Prm, EmpAy)
End Function

Function AyMapS1S2Ay(A, MapFunNm$) As S1S2()
Dim O() As S1S2, I
If Sz(A) > 0 Then
    For Each I In A
        PushObj O, S1S2(I, Run(MapFunNm, I))
    Next
End If
AyMapS1S2Ay = O
End Function

Function AyMapSy(A, MapFunNm$) As String()
AyMapSy = AyMapInto(A, MapFunNm, EmpSy)
End Function
Function AyMapAvSy(A, MapFunNm$, PrmAv) As String()
AyMapAvSy = AyMapAvInto(A, MapFunNm, PrmAv, EmpSy)
End Function
Function AyMapAvInto(A, MapFunNm$, PrmAv, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(A) > 0 Then
    Dim I
    Stop
    Dim Av(): Av = PrmAv: Av = AyIns(PrmAv)
    For Each I In A
        Asg I, Av(0)
        Push O, RunAv(MapFunNm, Av)
    Next
End If
AyMapAvInto = O
End Function

Function RunAv(FunNm$, Av())
Dim O
Select Case Sz(Av)
Case 0: O = Run(FunNm)
Case 1: O = Run(FunNm, Av(0))
Case 2: O = Run(FunNm, Av(0), Av(1))
Case 3: O = Run(FunNm, Av(0), Av(1), Av(2))
Case 4: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case Else: Stop
End Select
RunAv = O
End Function

Function AyMax(A)
Dim O: O = A(0)
Dim J&
For J = 1 To UB(A)
    O = Max(O, A(J))
Next
AyMax = O
End Function

Function AyMinus(A, B)
If Sz(B) = 0 Or Sz(A) = 0 Then AyMinus = A: Exit Function
Dim O: O = A: Erase O
Dim B1: B1 = B
Dim V
For Each V In A
    If Not AyHas(B, V) Then
        Push O, V
    End If
Next
AyMinus = O
End Function

Function AyMinusAp(A, ParamArray AyAp())
Dim O
If Sz(A) = 0 Then O = A: Erase O: GoTo X
O = A
Dim Av(): Av = AyAp
Dim Ay1, V
For Each Ay1 In Av
    O = AyMinus(O, A)
    If Sz(O) = 0 Then GoTo X
Next
X:
AyMinusAp = O
End Function

Function AyPair_Dic(A1, A2) As Dictionary
Dim N1&, N2&
N1 = Sz(A1)
N2 = Sz(A2)
If N1 <> N2 Then Stop
Dim O As New Dictionary
Dim J&
If Sz(A1) = 0 Then GoTo X
For J = 0 To N1 - 1
    O.Add A1(J), A2(J)
Next
X:
Set AyPair_Dic = O
End Function

Function AyRgH(Ay, At As Range) As Range
Set AyRgH = CellPutSq(At, AySqH(Ay))
End Function

Function AyRmvEle(Ay, Ele)
Dim Ix&: Ix = AyIx(Ay, Ele): If Ix = -1 Then AyRmvEle = Ay: Exit Function
AyRmvEle = AyRmvEleAt(Ay, AyIx(Ay, Ele))
End Function

Function AyRmvEleAt(Ay, Optional At&)
AyRmvEleAt = AyWhExclAtCnt(Ay, At)
End Function

Function AyRmvEmp(Ay)
If Sz(Ay) = 0 Then AyRmvEmp = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If Not IsEmp(I) Then Push O, I
Next
AyRmvEmp = O
End Function

Function AySqH(A) As Variant()
Dim O(), J&
ReDim O(1 To 1, 1 To Sz(A))
For J = 0 To UB(A)
    O(1, J + 1) = A(J)
Next
AySqH = O
End Function

Function AySqV(Ay) As Variant()
If Sz(Ay) = 0 Then Exit Function
Dim O(), R&
ReDim O(1 To Sz(Ay), 1 To 1)
R = 0
Dim V
For Each V In Ay
    R = R + 1
    O(R, 1) = V
Next
AySqV = O
End Function

Function AySrt(Ay, Optional Des As Boolean)
If Sz(Ay) = 0 Then AySrt = Ay: Exit Function
Dim Ix&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    O = AyIns(O, Ay(J), AySrt__Ix(O, Ay(J), Des))
Next
AySrt = O
End Function

Function AySrtInToIxAy__Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then AySrtInToIxAy__Ix& = O: Exit Function
        O = O + 1
    Next
    AySrtInToIxAy__Ix& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then AySrtInToIxAy__Ix& = O: Exit Function
    O = O + 1
Next
AySrtInToIxAy__Ix& = O
End Function

Function AySrtIntoIxAy(Ay, Optional Des As Boolean) As Long()
If Sz(Ay) = 0 Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = AyIns(O, J, AySrtInToIxAy__Ix(O, Ay, Ay(J), Des))
Next
AySrtIntoIxAy = O
End Function

Function AySrt__Ix&(A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In A
        If V > I Then AySrt__Ix = O: Exit Function
        O = O + 1
    Next
    AySrt__Ix = O
    Exit Function
End If
For Each I In A
    If V < I Then AySrt__Ix = O: Exit Function
    O = O + 1
Next
AySrt__Ix = O
End Function

Function AySy(A) As String()
If IsSy(A) Then AySy = A: Exit Function
Dim N&: N = Sz(A)
If N = 0 Then Exit Function
Dim I, O$(), J&
ReDim O(N - 1)
For J = 0 To N - 1
    O(J) = A(J)
Next
AySy = O
End Function

Function AyWs(A, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = SqRg(AySqV(A), NewA1).Parent
If Vis Then WsVis O
Set AyWs = O
End Function

Function AyWdt%(A)
Dim W%, I: If Sz(A) = 0 Then Exit Function
For Each I In A
    W = Max(Len(I), W)
Next
AyWdt = W
End Function

Function AyWhDist(A)
Dim O: O = A: Erase O
If Sz(A) > 0 Then
    Dim I
    For Each I In A
        PushNoDup O, I
    Next
End If
AyWhDist = O
End Function

Function AyWhDup(A)
Dim O
O = A
Erase O
If Sz(A) = 0 Then
    AyWhDup = O
    Exit Function
End If
Dim CntDry(): CntDry = AyCntDry(A)
Dim Dr
For Each Dr In CntDry
    If Dr(1) > 1 Then
        Push O, Dr(0)
    End If
Next
AyWhDup = O
End Function

Function AyWhExclAtCnt(Ay, At&, Optional Cnt& = 1)
If Cnt <= 0 Then AyWhExclAtCnt = Ay: Exit Function
Dim U&: U = UB(Ay)
If At > U Then Stop
If At < 0 Then Stop
If U = 0 Then AyWhExclAtCnt = Ay: Exit Function
Dim O: O = Ay
Dim J&
For J = At To U - Cnt
    O(J) = O(J + Cnt)
Next
ReDim Preserve O(U - Cnt)
AyWhExclAtCnt = O
End Function

Function AyWhExclNy0(A$(), ExclNy0) As String()
If IsMissing(ExclNy0) Then AyWhExclNy0 = A: Exit Function
Dim N$(): N = DftNy(ExclNy0)
AyWhExclNy0 = AyMinus(A, N)
End Function

Function AyWhFmTo(A, Fmix, Toix)
Dim O: O = A: Erase O
Dim J&
For J = Fmix To Toix
    Push O, A(J)
Next
AyWhFmTo = O
End Function
Function AyWhFTIx(A, X As FTIx)
AyWhFTIx = AyWhFmixToix(A, X.Fmix, X.Toix)
End Function
Function AyWhFmixToix(A, Fmix&, Toix&)
Dim O: O = A: Erase O
Dim J&
For J = Fmix To Toix
    Push O, A(J)
Next
AyWhFmixToix = O
End Function

Function AyWhPatn(A, Patn$, Optional ExclNy0) As String()
If Patn = "." And IsMissing(ExclNy0) Then
    AyWhPatn = AySy(A)
    Exit Function
End If
Dim I, O$(), R As RegExp
Set R = Re(Patn)
For Each I In A
    If R.Test(I) Then Push O, I
Next
AyWhPatn = AyWhExclNy0(O, ExclNy0)
End Function
Function AyWhPredNot(A, Pred$)
If Sz(A) = 0 Then AyWhPredNot = A: Exit Function
Dim O: O = A: Erase O
Dim J&
For J = 0 To UB(A)
    If Not Run(Pred, A(J)) Then
        Push O, A(J)
    End If
Next
AyWhPredNot = O
End Function

Function AyWhPred(A, Pred$)
If Sz(A) = 0 Then AyWhPred = A: Exit Function
Dim O: O = A: Erase O
Dim J&
For J = 0 To UB(A)
    If Run(Pred, A(J)) Then
        Push O, A(J)
    End If
Next
AyWhPred = O
End Function

Function AyWhSingleEle(A)
Dim O: O = A: Erase O
Dim CntDry(): CntDry = AyCntDry(A)
If Sz(CntDry) = 0 Then
    AyWhSingleEle = O
    Exit Function
End If
Dim Dr
For Each Dr In CntDry
    If Dr(1) = 1 Then
        Push O, Dr(0)
    End If
Next
AyWhSingleEle = O
End Function

Function Brk(A, Sep, Optional IsNoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then Stop
Dim S1$, S2$
    S1 = Left(A, P - 1)
    S2 = Mid(A, P + Len(Sep))
If Not IsNoTrim Then
    S1 = Trim(S1)
    S2 = Trim(S2)
End If
Set Brk = S1S2(S1, S2)
End Function

Function Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk1 = Brk1__X(A, P, Sep, NoTrim)
End Function

Function Brk1__X(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk1__X = S1S2(A, "")
    Else
        Set Brk1__X = S1S2(Trim(A), "")
    End If
    Exit Function
End If
Set Brk1__X = BrkAt(A, P, Sep, NoTrim)
End Function
Function BrkAt(A, P&, Sep, NoTrim As Boolean) As S1S2
Dim S1$, S2$
S1 = Left(A, P - 1)
S2 = Mid(A, P + Len(Sep))
If NoTrim Then
    Set BrkAt = S1S2(S1, S2)
Else
    Set BrkAt = S1S2(Trim(S1), Trim(S2))
End If
End Function

Function SqRg(A, At As Range) As Range
Dim O As Range: Set O = CellReSz(At, A)
O.Value = A
Set SqRg = O
End Function

Function CellPutSq(A As Range, Sq, Optional LoNm$) As ListObject
Set CellPutSq = RgLo(SqRg(Sq, A), LoNm)
End Function

Function CellReSz(A As Range, Sq) As Range
Set CellReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function

Function CmpTyAyOf_Cls_and_Std() As vbext_ComponentType()
Dim O(1) As vbext_ComponentType
O(0) = vbext_ct_ClassModule
O(1) = vbext_ct_StdModule
CmpTyAyOf_Cls_and_Std = O
End Function

Function CmpTy_Nm$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ct_ClassModule: O = "*Cls"
Case vbext_ct_StdModule: O = "*Std"
Case vbext_ct_Document: O = "*Doc"
Case Else: Stop
End Select
CmpTy_Nm = O
End Function

Function CollAddPfx(A As Collection, Pfx) As Collection
Dim O As New Collection, I
For Each I In A
    O.Add Pfx & I
Next
Set CollAddPfx = O
End Function

Function CurXls() As Excel.Application
Set CurXls = Excel.Application
End Function
Function CurWb() As Workbook
Set CurWb = CurXls.ActiveWorkbook
End Function

Function CurWs() As Worksheet
Set CurWs = CurXls.ActiveSheet
End Function

Function CurCdWin() As VBIDE.Window
Dim C As VBComponent: Set C = CurCmp: If IsNothing(C) Then Exit Function
Dim M As CodeModule: Set M = C.CodeModule: If IsNothing(M) Then Exit Function
Set CurCdWin = M.CodePane.Window
End Function

Function CurCmp() As VBComponent
Set CurCmp = CurMd.Parent
End Function

Function CurFunDNm$()
Dim M$: M = CurMthNm
If M = "" Then Exit Function
If Not MdIsStd(CurMd) Then Exit Function
CurFunDNm = CurMdDNm & "." & M
End Function
Function XXX()

End Function
Function CurSrc() As String()
CurSrc = MdSrc(CurMd)
End Function
Function CurMd() As CodeModule
Set CurMd = CurVbe.ActiveCodePane.CodeModule
End Function

Function CurMdDNm$()
CurMdDNm = MdDNm(CurMd)
End Function

Function CurMdNm$()
CurMdNm = CurCmp.Name
End Function

Function CurMth() As Mth
Dim Nm$: Nm = CurMthNm
If Nm = "" Then Stop
Set CurMth = Mth(CurMd, Nm)
End Function

Function CurMthDNm$()
CurMthDNm = CurMdDNm & "." & CurMthNm
End Function

Function CurMthNm$()
Dim L1&, L2&, C1&, C2&, K As vbext_ProcKind
Dim O$
With CurVbe.ActiveCodePane
    On Error GoTo X
    .GetSelection L1, C1, L2, C2
    On Error GoTo 0
    O = .CodeModule.ProcOfLine(L1, K)
End With
If O = "" Then Stop
CurMthNm = O
Exit Function
X:
End Function

Function CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Function

Function CurPjNm$()
CurPjNm = CurPj.Name
End Function

Function CurPjPth$()
CurPjPth = PjPth(CurPj)
End Function

Function CurVbe() As Vbe
Set CurVbe = CurXls.Vbe
End Function

Function CvFTNo(A) As FTNo
Set CvFTNo = A
End Function

Function CvFTIx(A) As FTIx
Set CvFTIx = A
End Function

Function CvMd(A) As CodeModule
Set CvMd = A
End Function
Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function
Function CvS1S2(A) As S1S2
Set CvS1S2 = A
End Function
Function CvMth(A) As Mth
Set CvMth = A
End Function

Function CvPj(I) As VBProject
Set CvPj = I
End Function

Function CvSy(A) As String()
CvSy = A
End Function

Function DCRsltBrw(A As DCRslt)

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

Function DCRsltLy(A As DCRslt) As String()
With A
Dim A1$(): A1 = DCRsltLy__AExcess(.AExcess)
Dim A2$(): A2 = DCRsltLy__BExcess(.BExcess)
Dim A3$(): A3 = DCRsltLy__Dif(.ADif, .BDif)
Dim A4$(): A4 = DCRsltLy__Sam(.Sam)
End With
DCRsltLy = AyAddAp(A1, A2, A3, A4)
End Function

Function DCRsltLy__AExcess(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, Ly$(), S1$, S2$, S(0) As S1S2
S2 = "!" & "Er AExcess"
For Each K In A.Keys
    S1 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2Ay_FmtLy(S)
    PushAy O, Ly
Next
DCRsltLy__AExcess = O
End Function

Function DCRsltLy__BExcess(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, Ly$(), S1$, S2$, S(0) As S1S2
S1 = "!" & "Er BExcess"
For Each K In A.Keys
    S2 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2Ay_FmtLy(S)
    PushAy O, Ly
Next
DCRsltLy__BExcess = O
End Function

Function DCRsltLy__Dif(A As Dictionary, B As Dictionary) As String()
If A.Count <> B.Count Then Stop
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S(0) As S1S2, Ly$()
For Each K In A
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & B(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2Ay_FmtLy(S)
    PushAy O, Ly
Next
DCRsltLy__Dif = O
End Function

Function DCRsltLy__Sam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S1S2
For Each K In A.Keys
    PushObj S, S1S2("*Same", K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K))
Next
DCRsltLy__Sam = S1S2Ay_FmtLy(S)
End Function

Function DDNmThird$(A)
Dim Ay$(): Ay = Split(A, "."): If Sz(Ay) <> 3 Then Stop
DDNmThird = Ay(2)
End Function

Function DftMd(MdDNm0$)
If MdDNm0 = "" Then
    Set DftMd = CurMd
Else
    Set DftMd = Md(MdDNm0)
End If
End Function

Function DftMdDNm$(MdDNm0$)
If MdDNm0 = "" Then
    DftMdDNm = CurMdNm
Else
    DftMdDNm = MdDNm0
End If
End Function

Function DftMdySy(A$) As String()
DftMdySy = DftNy(A)
End Function

Function DftMth(MthDNm0$) As Mth
If MthDNm0 = "" Then
    Set DftMth = CurMth
    Exit Function
End If
Set DftMth = MthDNm_Mth(MthDNm0)
End Function

Function DftMthNm$(MthNm0$)
If MthNm0 = "" Then
    DftMthNm = CurMthNm
    Exit Function
End If
DftMthNm = MthNm0
End Function

Function DftNy(Ny0) As String()
Dim T As VbVarType: T = VarType(Ny0)
If T = vbEmpty Then Exit Function
If IsMissing(Ny0) Then Exit Function
If T = vbString Then
    DftNy = SplitSsl(Ny0)
    Exit Function
End If
DftNy = Ny0
End Function

Function DftPj(PjNm0$)
If PjNm0 = "" Then
    Set DftPj = CurPj
Else
    Set DftPj = Pj(PjNm0)
End If
End Function

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

Function DrsWs(A As Drs) As Worksheet
Dim O As Worksheet, R As Range
Set O = NewWs
AyRgH A.Fny, WsA1(O)
Set R = CellPutSq(WsRC(O, 2, 1), DrySq(A.Dry))
Set DrsWs = O
End Function

Function DryNCol&(A())
Dim O&, Dr
For Each Dr In A
    O = Max(O, Sz(Dr))
Next
DryNCol = O
End Function
Function DryFny_Sq(A() As Variant, Fny$()) As Variant()
Dim NCol&, NRow&
    NCol = Max(DryNCol(A), Sz(Fny))
    NRow = Sz(A)
Dim O()
ReDim O(1 To 1 + NRow, 1 To NCol)
Dim C&, R&, Dr()
    For C = 1 To Sz(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NRow
        Dr = A(R - 1)
        For C = 1 To Min(Sz(Dr), NCol)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
DryFny_Sq = O
End Function
Function DrySq(A() As Variant) As Variant()
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

Function DupFunFNyGpAy_AllSameCnt%(A)
If Sz(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupFunFNyGp_IsDup(Gp) Then O = O + 1
Next
DupFunFNyGpAy_AllSameCnt = O
End Function

Function DupFunFNyGp_Dry(Ny$()) As Variant()
'Given Ny: Each Nm in Ny is FunNm:PjNm.MdNm
'          It has at least 2 ele
'          Each FunNm is same
'Return: N-Dr of Fields {Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src}
'        where N = Sz(Ny)-1
'        where each-field-(*-1)-of-Dr comes from Ny(0)
'        where each-field-(*-2)-of-Dr comes from Ny(1..)

Dim Md1$, Pj1$, Nm$
    FunFNm_BrkAsg Ny(0), Nm, Pj1, Md1
Dim Mth1 As New Mth
    Mth1.Nm = Nm
    Set Mth1.Md = Md(Pj1 & "." & Md1)
Dim Src1$
    Src1 = MthLines(Mth1)
Dim Mdy1$, Ty1$
    MthBrkAsg Mth1, Mdy1, Ty1
Dim O()
    Dim J%
    For J = 1 To UB(Ny)
        Dim Pj2$, Nm2$, Md2$
            FunFNm_BrkAsg Ny(J), Nm2, Pj2, Md2: If Nm2 <> Nm Then Stop
        Dim Mth2 As New Mth
            Mth2.Nm = Nm
            Set Mth2.Md = Md(Pj2 & "." & Md2)
        Dim Src2$
            Src2 = MthLines(Mth2)
        Dim Mdy2$, Ty2$
            MthBrkAsg Mth2, Mdy2, Ty2

        Push O, Array(Nm, _
                    Mdy1, Ty1, Pj1, Md1, _
                    Mdy2, Ty2, Pj2, Md2, Src1, Src2, Pj1 = Pj2, Md1 = Md2, Src1 = Src2)
    Next
DupFunFNyGp_Dry = O
End Function

Function DupFunFNyGp_IsDup(Ny) As Boolean
DupFunFNyGp_IsDup = AyIsAllEleEq(AyMap(Ny, "FunFNm_MthLines"))
End Function

Function DupFunFNy_GpAy(A$()) As Variant()
Dim O(), J%, M$()
Dim L$ ' LasMthNm
L = Brk(A(0), ":").S1
Push M, A(0)
Dim B As S1S2
For J = 1 To UB(A)
    Set B = Brk(A(J), ":")
    If L <> B.S1 Then
        Push O, M
        Erase M
        L = B.S1
    End If
    Push M, A(J)
Next
If Sz(M) > 0 Then
    Push O, M
End If
DupFunFNy_GpAy = O
End Function

Function DupFunFNy_SamMthBdyFunFNy(A$(), Vbe As Vbe) As String()
Dim Gp(): Gp = DupFunFNy_GpAy(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupFunFNyGp_IsDup(Ny) Then
        For Each N In Ny
            Push O, N
        Next
    End If
Next
DupFunFNy_SamMthBdyFunFNy = O
End Function

Function DupMthFNyGp_CmpLy(A, Optional OIx% = -1, Optional OSam% = -1, Optional InclSam As Boolean) As String()
'DupMthFNyGp is Variant/String()-of-MthFNm with all mth-nm is same
'MthFNm is MthNm in FNm-fmt
'          Mth is Prp/Sub/Fun in Md-or-Cls
'          FNm-fmt which is 'Nm:Pj.Md'
'DupMthFNm is 2 or more MthFNy with same MthNm
Ass DupMthFNyGp_IsVdt(A)
Dim J%, I%
Dim LinesAy$()
Dim UniqLinesAy$()
    LinesAy = AyMapSy(A, "FunFNm_MthLines")
    UniqLinesAy = AyWhDist(LinesAy)
Dim MthNm$: MthNm = Brk(A(0), ":").S1
Dim Hdr$(): Hdr = DupMthFNyGp_CmpLy__1Hdr(OIx, MthNm, Sz(A))
Dim Sam$(): Sam = DupMthFNyGp_CmpLy__2Sam(InclSam, OSam, A, LinesAy)
Dim Syn$(): Syn = DupMthFNyGp_CmpLy__3Syn(UniqLinesAy, LinesAy, A)
Dim Cmp$(): Cmp = DupMthFNyGp_CmpLy__4Cmp(UniqLinesAy, LinesAy, A)
DupMthFNyGp_CmpLy = AyAddAp(Hdr, Sam, Syn, Cmp)
End Function

Function DupMthFNyGp_IsVdt(A) As Boolean
If Not IsSy(A) Then Exit Function
If Sz(A) <= 1 Then Exit Function
Dim N$: N = Brk(A(0), ":").S1
Dim J%
For J = 1 To UB(A)
    If N <> Brk(A(J), ":").S1 Then Exit Function
Next
DupMthFNyGp_IsVdt = True
End Function

Function EitherL(A) As Either
Asg A, EitherL.Left
EitherL.IsLeft = True
End Function

Function EitherR(A) As Either
Asg A, EitherR.Right
End Function
Function EmpMdAy() As CodeModule
End Function
Function EmpAy() As Variant()
End Function

Function EmpIntAy() As Integer()
End Function

Function EmpRfAy() As Reference()
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
Function CvDic(A) As Dictionary
Set CvDic = A
End Function
Sub DicPush(O As Dictionary, M As Dictionary)
'Assume there is no duplicated key
If M.Count = 0 Then Exit Sub
Dim K
For Each K In M.Keys
    O.Add K, M(K)
Next
End Sub
Function RmvUSfx$(A)
Dim J%, Fnd As Boolean, C%
For J = Len(A) To 2 Step -1 ' don't find the first char if non-UCase, to use 'To 2'
    C = Asc(Mid(A, J, 1))
    If Not AscIsUCase(C) Then
        Fnd = True
        Exit For
    End If
Next
If Fnd Then
    RmvUSfx = Left(A, J)
Else
    RmvUSfx = A
End If
End Function
Function DicIsEmp(A As Dictionary) As Boolean
DicIsEmp = A.Count = 0
End Function

Function EmpDicAy() As Dictionary()
End Function
Function DicMap(A As Dictionary, ValMapFun$) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, Run(ValMapFun, A(K))
Next
Set DicMap = O
End Function
Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function
Function EmpSy() As String()
End Function

'Sub FfnDlt(Ffn)
'If FfnIsExist(Ffn) Then Kill Ffn
'End Sub
'Function FfnExt$(Ffn)
'Dim P%: P = InStrRev(Ffn, ".")
'If P = 0 Then Exit Function
'FfnExt = Mid(Ffn, P)
'End Function
'Function FfnFdr$(Ffn)
'FfnFdr = PthFdr(FfnPth(Ffn))
'End Function
Function FfnFn$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then FfnFn = A: Exit Function
FfnFn = Mid(A, P + 1)
End Function

Function FfnFnn$(A)
FfnFnn = FfnRmvExt(FfnFn(A))
End Function

Function FfnIsExist(A) As Boolean
FfnIsExist = Fso.FileExists(A)
End Function

Function FfnPth$(A)
Dim P%: P = InStrRev(A, "\")
If P = 0 Then Exit Function
FfnPth = Left(A, P)
End Function

Function FfnRmvExt$(A)
Dim P%: P = InStrRev(A, ".")
If P = 0 Then FfnRmvExt = Left(A, P): Exit Function
FfnRmvExt = Left(A, P - 1)
End Function

Function FTNoAy_LinCnt%(A() As FTNo)
Dim O%, M
For Each M In A
    O = O + FTNo_LinCnt(CvFTNo(M))
Next
End Function

Function FTNo_LinCnt%(A As FTNo)
Dim O%
O = A.Tono - A.Fmno + 1
If O < 0 Then Stop
FTNo_LinCnt = O
End Function

Function FTIx_FTNo(A As FTIx) As FTNo
Set FTIx_FTNo = FTNo(A.Fmix + 1, A.Toix + 1)
End Function

Function FTIx_LinCnt%(A As FTIx)
Dim O%
O = A.Toix - A.Fmix + 1
If O < 0 Then Stop
FTIx_LinCnt = O
End Function

Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim O$: O = Replace(QQVbl, "|", vbCrLf)
Dim Av(): Av = Ap
Dim I
For Each I In Av
    O = Replace(O, "?", I, Count:=1)
Next
FmtQQ = O
End Function

Function FnyOf_MthKey() As String()
FnyOf_MthKey = SslSy("PjNm MdNm Priority Nm Ty Mdy")
End Function

Function Fso() As FileSystemObject
Set Fso = New FileSystemObject
End Function

Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function FunFNm_MdDNm$(A)
FunFNm_MdDNm = Brk(A, ":").S2
End Function

Function FunFNm_MthLines$(A)
FunFNm_MthLines = MthLines(MthFNm_Mth(A))
End Function

Function FunFNy_DupFunFNy(A$(), Optional IsSamMthBdyOnly As Boolean) As String()
If Sz(A) = 0 Then Exit Function
Dim A1$(): A1 = AySrt(A)
Dim O$(), M$(), J&, Nm$
Dim L$ ' LasFunNm
L = MthFNm_Nm(A1(0))
Push M, A1(0)
For J = 1 To UB(A1)
    Nm = MthFNm_Nm(A1(J))
    If L = Nm Then
        Push M, A1(J)
    Else
        L = Nm
        If Sz(M) = 1 Then
            M(0) = A1(J)
        Else
            PushAy O, M
            Erase M
        End If
    End If
Next
If Sz(M) > 1 Then
    PushAy O, M
End If
FunFNy_DupFunFNy = O
End Function

Function FunNm_CmpLy(A, Optional InclSam As Boolean) As String()
'Found all Fun with given name and compare if it is same
'Note: Fun is Fun/Sub/Prp-in-Md
Dim O$()
Dim N$(): N = FunNm_DupFunFNy(A)
DupFunFNy_ShwNotDupMsg N, A
If Sz(N) <= 1 Then Exit Function
FunNm_CmpLy = DupMthFNyGp_CmpLy(N, InclSam:=InclSam)
End Function

Function FunNm_DupFunFNy(A) As String()
FunNm_DupFunFNy = VbeFunFNy(CurVbe, FunNmPatn:="^" & A & "$", ExclFunNy0:="ZZZ__Tst", Mdy0:="Public")
End Function
Private Sub ZZ_MthNm_MthPfx()
Debug.Assert MthNm_MthPfx("Add_Cls") = "Add"
End Sub
Private Sub ZZ_MthNm_MthPfx__BrwAll()
Dim Ay$(): Ay = VbeMthNy(CurVbe)
Dim Ay1$(): Ay1 = AyMapSy(Ay, "MthNm_MthPfx")
WsVis AyAB_Ws(Ay, Ay1)
End Sub
Function AyAB_Ws(A, B) As Worksheet
Dim N&: N = Sz(A)
If N <> Sz(B) Then Stop
Dim Ws As Worksheet: Set Ws = NewWs
WsRC(Ws, 1, 1).Value = "A"
WsRC(Ws, 1, 2).Value = "B"
CellPutAyV WsRC(Ws, 2, 1), A
CellPutAyV WsRC(Ws, 2, 2), B
RgLo WsRCRC(Ws, 1, 1, N + 1, 2)
Set AyAB_Ws = Ws
End Function
Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function
Function CellPutAyH(A As Range, AyH) As Range
Set CellPutAyH = CellPutSq(A, AySqH(AyH))
End Function
Function CellPutAyV(A As Range, AyV) As Range
Set CellPutAyV = CellPutSq(A, AySqV(AyV))
End Function
Sub ZZZ_RmvPfxAy()
Const A1$ = "ZZZ_ABC"
Const A2$ = "ZZ_ABC"
Const C$ = "ZZ_|ZZZ_"
Debug.Assert RmvPfxAy(A1, C) = "ABC"
Debug.Assert RmvPfxAy(A2, C) = "ABC"
End Sub
Function RmvPfxAy$(A, PfxAyVbl$)
Dim P$(): P = SplitVBar(PfxAyVbl)
Dim Pfx
For Each Pfx In P
    If HasPfx(A, Pfx) Then
        RmvPfxAy = RmvPfx(A, Pfx)
        Exit Function
    End If
Next
RmvPfxAy = A
End Function
Function MthNm_MthPfx$(A)
Dim A0$
    A0 = Brk1(RmvPfxAy(A, "ZZ_|ZZZ_"), "__").S1
With Brk2(A0, "_")
    If .S1 <> "" Then
        MthNm_MthPfx = .S1
        Exit Function
    End If
End With
Dim P2%
Dim Fnd As Boolean
    Dim C%
    Fnd = False
    For P2 = 2 To Len(A0)
        C = Asc(Mid(A0, P2, 1))
        If AscIsLCase(C) Then Fnd = True: Exit For
    Next
'---
    If Not Fnd Then Exit Function
Dim P3%
Fnd = False
    For P3 = P2 + 1 To Len(A0)
        C = Asc(Mid(A0, P3, 1))
        If AscIsUCase(C) Or AscIsDigit(C) Then Fnd = True: Exit For
    Next
'--
If Fnd Then
    MthNm_MthPfx = Left(A0, P3 - 1)
    Exit Function
End If
MthNm_MthPfx = A
End Function

Function MthNm_ProperMdNm$(A)
If A = "ZZZ__Tst" Then Exit Function
Dim P$: P = MthNm_MthPfx(A): If P = "" Then Exit Function
MthNm_ProperMdNm = "M_" & P
End Function

Function FxWb(A) As Workbook
Set FxWb = Xls.Workbooks.Open(A)
End Function

Function FxaNm_Fxa$(A)
FxaNm_Fxa = CurPjPth & A & ".xlam"
End Function

Function HasPfx(S, Pfx) As Boolean
HasPfx = Left(S, Len(Pfx)) = Pfx
End Function

Function HasSubStr(A, SubStr$) As Boolean
HasSubStr = InStr(A, SubStr) > 0
End Function

Function IntAy_Add1(A%()) As Integer()
IntAy_Add1 = IntAy_AddN(A, 1)
End Function

Function IntAy_AddN(A%(), N%) As Integer()
If Sz(A) = 0 Then Exit Function
Dim O%(), U&
U = UB(A)
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = A(J) + N
Next
IntAy_AddN = O
End Function

Function IsDigit(A) As Boolean
IsDigit = "0" <= A And A <= "9"
End Function

Function IsEmp(V) As Boolean
IsEmp = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsEmpty(V) Then Exit Function
If IsStr(V) Then
   If V = "" Then Exit Function
End If
If IsArray(V) Then
   If Sz(V) = 0 Then Exit Function
End If
IsEmp = False
End Function

Function IsFun(A As Mth) As Boolean
If Not MdIsStd(A.Md) Then Exit Function
IsFun = True
End Function

Function IsLetter(A) As Boolean
Dim C1$: C1 = UCase(A)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsMdNm(A) As Boolean
Select Case Left(A, 2)
Case "M_", "S_", "F_", "G_"
    IsMdNm = True
End Select
End Function

Function IsMthTy(A$) As Boolean
Select Case A
Case "Function", "Property Let", "Property Set", "Sub", "Function": IsMthTy = True
End Select
End Function

Function IsNm(A) As Boolean
If Not IsLetter(FstChr(A)) Then Exit Function
Dim L%: L = Len(A)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(A, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsNmChr(A$) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Function
If A = "_" Then Exit Function
If IsDigit(A) Then Exit Function
IsNmChr = False
End Function

Function IsNothing(A) As Boolean
IsNothing = TypeName(A) = "Nothing"
End Function

Function IsPfx(A, Pfx) As Boolean
IsPfx = Left(A, Len(Pfx)) = Pfx
End Function

Function IsPrim(A) As Boolean
Select Case VarType(A)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   IsPrim = True
End Select
End Function

Function IsPun(C) As Boolean
If IsLetter(C) Then Exit Function
If IsDigit(C) Then Exit Function
If C = "_" Then Exit Function
IsPun = True
End Function

Function IsStr(A) As Boolean
IsStr = VarType(A) = vbString
End Function

Function IsSy(A) As Boolean
IsSy = VarType(A) = vbArray + vbString
End Function

Function ItrAy(A, OIntoAy)
Dim O: O = OIntoAy: Erase O
Dim I
For Each I In A
    Push O, I
Next
ItrAy = O
End Function

Function ItrNy(Itr, Optional Patn$ = ".", Optional ExclNy0) As String()
Dim I, O$()
For Each I In Itr
    Push O, CallByName(I, "Name", VbGet)
Next
ItrNy = AyWhPatn(O, Patn, ExclNy0)
End Function

Function JnComma$(A)
JnComma = Join(A, ",")
End Function

Function JnCrLf$(A)
JnCrLf = Join(A, vbCrLf)
End Function

Function LasChr$(A)
LasChr = Right(A, 1)
End Function

Function LinIsCd(A) As Boolean
Dim L$: L = Trim(A)
If A = "" Then Exit Function
If Left(A, 1) = "'" Then Exit Function
LinIsCd = True
End Function

Function LinIsMthLin(A) As Boolean
LinIsMthLin = AyHas(SyOf_PrpSubFun, LinT1(LinRmvMdy(A)))
End Function

Function LinIsTstSub(L$) As Boolean
LinIsTstSub = True
If IsPfx(L, "Sub Tst()") Then Exit Function
If IsPfx(L, "Sub Tst()") Then Exit Function
If IsPfx(L, "Friend Sub Tst()") Then Exit Function
If IsPfx(L, "Sub ZZZ__Tst()") Then Exit Function
If IsPfx(L, "Sub ZZZ__Tst()") Then Exit Function
If IsPfx(L, "Friend Sub ZZZ__Tst()") Then Exit Function
LinIsTstSub = False
End Function

Function LinLCCOpt(A, MthNm$, Lno%) As LCCOpt
Dim M$: M = LinMthNm(A)
If M <> MthNm Then Set LinLCCOpt = New LCCOpt: Exit Function
Dim C1%, C2%
C1 = InStr(A, MthNm)
C2 = C1 + Len(MthNm)
Set LinLCCOpt = LCCOpt(LCC(Lno, C1, C2))
End Function

Function LinMdy$(A)
LinMdy = LinPfxOfAy(A, SyOf_Mdy)
End Function

Function LinMthNm$(A)
Dim L$: L = LinRmvMdy(A)
Dim B$: B = LinShiftMthTy(L): If B = "" Then Exit Function
LinMthNm = LinNm(L)
End Function

Function LinMthTy$(A)
Dim A1$, A2$
A1 = LinRmvMdy(A)
A2 = LinT1(A1)
LinMthTy = LinPfxOfAy(A2, SyOf_MthTy)
End Function

Function LinNm$(A)
Dim J%
If Not IsLetter(Left(A, 1)) Then Exit Function
For J = 2 To Len(A)
    If Not IsNmChr(Mid(A, J, 1)) Then
        LinNm = Left(A, J - 1)
        Exit Function
    End If
Next
LinNm = A
End Function

Function LinPfxOfAy$(A, PfxAy$())
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, Pfx) Then LinPfxOfAy = Pfx: Exit Function
Next
End Function

Function LinPrpSubFun$(A)
LinPrpSubFun = LinPfxOfAy(LinRmvMdy(A), SyOf_PrpSubFun)
End Function

Function LinRmvMdy$(A)
LinRmvMdy = LinRmvPfxOfAyAndTrim(A, SyOf_Mdy)
End Function

Function LinRmvPfxOfAyAndTrim$(A, PfxAy$())
Dim L$: L = A
LinShiftPfxAyAndLTrim L, PfxAy
LinRmvPfxOfAyAndTrim = L
End Function

Function LinRmvT1$(A)
Dim O$: O = A
LinShiftT1 O
LinRmvT1 = O
End Function

Function LinShiftBktEnclosedStr$(O$)
If FstChr(O) <> "(" Then Stop
Dim J%
Dim Fnd As Boolean
    Dim Cnt%
    For J = 2 To Len(O)
        Select Case Mid(O, J, 1)
        Case ")"
            If Cnt = 0 Then
                Fnd = True
                Exit For
            Else
                Cnt = Cnt - 1
                If Cnt < 0 Then Stop
            End If
        Case "("
            Cnt = Cnt + 1
        End Select
    Next
If Not Fnd Then Stop
LinShiftBktEnclosedStr = Left(O, J)
O = Mid(O, J + 1)

End Function

Function LinShiftMdy$(O$)
LinShiftMdy = LinShiftPfxAyAndLTrim(O, SyOf_Mdy)
End Function
Function LinShiftShtMdy$(O$)
LinShiftShtMdy = MdyShtMdy(LinShiftPfxAyAndLTrim(O, SyOf_Mdy))
End Function

Function LinShiftMthTy$(O$)
LinShiftMthTy = LinShiftPfxAyAndLTrim(O, SyOf_MthTy)
End Function
Function LinShiftMthShtTy$(O$)
LinShiftMthShtTy = MthTy_MthShtTy(LinShiftPfxAyAndLTrim(O, SyOf_MthTy))
End Function

Function LinShiftNm$(O$)
Dim A$: A = LTrim(O)
Dim Nm$: Nm = LinNm(A): If Nm = "" Then Exit Function
LinShiftNm = Nm
O = RmvPfx(A, Nm)
End Function

Function LinShiftPfxAyAndLTrim$(O$, PfxAy$())
Dim A$: A = LTrim(O)
Dim Pfx$: Pfx = LinPfxOfAy(A, PfxAy)
If Pfx <> "" Then
    O = LTrim(RmvPfx(A, Pfx))
    LinShiftPfxAyAndLTrim = Pfx
End If
End Function

Function LinShiftT1$(O$)
With Brk1(LTrim(O), " ")
    LinShiftT1 = .S1
    O = .S2
End With
End Function

Function LinShiftTySfxChr$(O$)
Dim F$: F = FstChr(O)
If InStr("#!@#$%^&", F) > 0 Then
    LinShiftTySfxChr = F
    O = RmvFstChr(O)
End If
End Function

Function LinT1$(L)
Dim A$: A = LTrim(L)
Dim P%: P = InStr(A, " ")
If P = 0 Then LinT1 = RTrim(A): Exit Function
LinT1 = Left(A, P - 1)
End Function

Function LinesAy_FmtLy(A$()) As String()
Dim LyAy()
    LyAy = AyMap(A, "SplitCrLf")
Dim W%()
    W = AyMapInto(LyAy, "AyWdt", EmpIntAy)
Dim NRowAy%()
    NRowAy = AyMapInto(LyAy, "Sz", EmpIntAy)
Dim NRow%
    NRow = AyMax(NRowAy)
Dim O$()
    Dim J%, Hdr$
    Hdr = WdtAy_HdrLin(W)
    Push O, Hdr
    For J = 0 To NRow - 1
        Push O, LyAy_Lin(LyAy, W, J)
    Next
    Push O, Hdr
LinesAy_FmtLy = O
End Function

Function LinesAy_Wdt%(A)
If Sz(A) = 0 Then Exit Function
Dim O%, J&, M%, L
For Each L In A
   O = Max(O, LinesWdt(L))
Next
LinesAy_Wdt = O
End Function

Function LinesBoxLy(A) As String()
LinesBoxLy = LyBoxLy(SplitCrLf(A))
End Function

Function LinesLinCnt%(A$)
LinesLinCnt = StrSubStrCnt(A, vbCrLf) + 1
End Function

Function LinesSqV(Lines$) As Variant
LinesSqV = AySqV(SplitCrLf(Lines))
End Function

Function LinesTrimEnd$(A$)
LinesTrimEnd = Join(LyTrimEnd(SplitCrLf(A)), vbCrLf)
End Function

Function LinesUnderLin$(Lines)
LinesUnderLin = StrDup("-", LinesWdt(Lines))
End Function

Function LinesVbl$(A)
LinesVbl = Replace(A, vbCrLf, "|")
End Function

Function LinesWdt%(A)
LinesWdt = AyWdt(SplitCrLf(A))
End Function

Function LoQt(A As ListObject) As QueryTable
On Error Resume Next
Set LoQt = A.QueryTable
End Function

Function LyAy_Lin$(A(), WdtAy%(), Ix%)
Dim J%, W%, I$, Ly$(), Dr$()
For J = 0 To UB(A)
    Ly = A(J)
    W% = WdtAy(J)
    If UB(Ly) >= Ix Then
        I = Ly(Ix)
    Else
        I = ""
    End If
    Push Dr, AlignL(I, W)
Next
LyAy_Lin = "| " + Join(Dr, " | ") + " |"
End Function

Function LyBoxLy(A) As String()
If Sz(A) = 0 Then Exit Function
Dim W%: W = AyWdt(A)
Dim H$: H = "|" & StrDup("-", W + 2) & "|"
Dim O$()
Push O, H
Dim I
For Each I In A
    Push O, "| " & AlignL(I, W) + " |"
Next
Push O, H
LyBoxLy = O
End Function

Function LyTrimEnd(Ly) As String()
If Sz(Ly) = 0 Then Exit Function
Dim L$
Dim J&
For J = UB(Ly) To 0 Step -1
    L = Trim(Ly(J))
    If Trim(Ly(J)) <> "" Then
        Dim O$()
        O = Ly
        ReDim Preserve O(J)
        LyTrimEnd = O
        Exit Function
    End If
Next
End Function

Function Max(A, B)
If A > B Then
    Max = A
Else
    Max = B
End If
End Function

Function MaxCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(Application.Version = "16.0", 16384, 255)
End If
MaxCol = C
End Function

Function MaxRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(Application.Version = "16.0", 1048576, 65535)
End If
MaxRow = R
End Function

Function Md(MdDNm) As CodeModule
Dim A$: A = MdDNm
Dim P As VBProject
Dim MdNm$
    Dim L%
    L = InStr(A, ".")
    If L = 0 Then
        Set P = CurPj
        MdNm = A
    Else
        Dim PjNm$
        PjNm = Left(A, L - 1)
        Set P = Pj(PjNm)
        MdNm = Mid(A, L + 1)
    End If
Set Md = P.VBComponents(MdNm).CodeModule
End Function

Function MdMthLinAy(A As CodeModule) As String()
MdMthLinAy = SrcMthLinAy(MdSrc(A))
End Function

Function MdCmpTy(A As CodeModule) As vbext_ComponentType
MdCmpTy = A.Parent.Type
End Function

Function MdDNm$(A As CodeModule)
MdDNm = MdPjNm(A) & "." & MdNm(A)
End Function

Function MdDic(A As CodeModule, Optional ExclDcl As Boolean) As Dictionary
Set MdDic = SrcDicOfMthNmzzzMthLines(MdSrc(A), ExclDcl)
End Function

Function MdDicOfMthKeyzzzMthLines(A As CodeModule) As Dictionary
Set MdDicOfMthKeyzzzMthLines = SrcDicOfMthKeyzzzMthLines(MdSrc(A), MdPjNm(A), MdNm(A))
End Function

Function MdDicOfMthNmzzzMthLines(A As CodeModule) As Dictionary
Set MdDicOfMthNmzzzMthLines = SrcDicOfMthNmzzzMthLines(MdSrc(A))
End Function

Function MdFunFNy(A As CodeModule, Optional FunNmPatn$ = ".", Optional ExclFunNy0$, Optional Mdy0$, Optional MthTy0$) As String()
Dim P$, M$, Sfx$, S$(), N$()
    P = MdPjNm(A)
    M = MdNm(A)
    Sfx = ":" & P & "." & M
    S = MdSrc(A)
    N = SrcMthNy(S, MthNmPatn:=FunNmPatn, ExclMthNy0:=ExclFunNy0, Mdy0:=Mdy0$)
MdFunFNy = AyAddSfx(N, Sfx)
End Function

Function MdFunPfxAy(A As CodeModule) As String()
Dim O$(), N, Ay$()
Ay = MdMthNy(A, IsNoMdNmPfx:=True, Mdy0:="Public")
If Sz(Ay) = 0 Then Exit Function
For Each N In Ay
    PushNoDup O, MthNm_MthPfx(N)
Next
MdFunPfxAy = O
End Function

Function MdHasMth(A As CodeModule, MthNm$) As Boolean
MdHasMth = MdMthFmno(A, MthNm) > 0
End Function

Function MdHasTstSub(A As CodeModule) As Boolean
Dim I
For Each I In MdLy(A)
    If I = "Friend Sub ZZZ__Tst()" Then MdHasTstSub = True: Exit Function
    If I = "Sub ZZZ__Tst()" Then MdHasTstSub = True: Exit Function
Next
End Function

Function MdIsAllRemarked(Md As CodeModule) As Boolean
Dim J%, L$
For J = 1 To Md.CountOfLines
    If Left(Md.Lines(J, 1), 1) <> "'" Then Exit Function
Next
MdIsAllRemarked = True
End Function

Function MdIsCls(A As CodeModule) As Boolean
MdIsCls = A.Parent.Type = vbext_ct_ClassModule
End Function

Function MdIsFunGpMd(A As CodeModule) As Boolean
'A Md is a FunGpMd must be with Name begins with M_ or S_
'so that all public-function or ZZ_-function has Fun-ProperMdNm matches with its module-name
If A.Parent.Type <> vbext_ct_StdModule Then Exit Function
Dim MdN$: MdN = MdNm(A)
    Dim Pfx$
    Pfx = Left(MdN, 2)
MdIsFunGpMd = Pfx = "M_" Or Pfx = "S_"
End Function

Function MdIsStd(A As CodeModule) As Boolean
MdIsStd = A.Parent.Type = vbext_ct_StdModule
End Function

Function MdIsStdMd(A As CodeModule) As Boolean
MdIsStdMd = A.Parent.Type = vbext_ct_StdModule
End Function

Function MdLines$(A As CodeModule)
With A
    If .CountOfLines = 0 Then Exit Function
    MdLines = .Lines(1, .CountOfLines)
End With
End Function

Function MdLy(A As CodeModule) As String()
MdLy = Split(MdLines(A), vbCrLf)
End Function

Function MdMthAy(A As CodeModule, Optional MthNmPatn$ = ".", Optional Mdy0$) As Mth()
Dim N$(), J%, O() As Mth
N = MdMthNy(A, MthNmPatn, IsNoMdNmPfx:=True, Mdy0:=Mdy0)
Dim U%: U = UB(N)
If U >= 0 Then
    ReDim O(U)
    For J = 0 To U
        Set O(J) = Mth(A, N(J))
    Next
End If
MdMthAy = O
End Function

Function MdMthFmno(A As CodeModule, MthNm$)
MdMthFmno = SrcMthFmix(MdSrc(A), MthNm) + 1
End Function

Function MdMthSq(A As CodeModule) As Variant()
MdMthSq = MthKy_Sq(MdMthKy(A, True))
End Function

Function PjMthSq(A As VBProject) As Variant()
PjMthSq = MthKy_Sq(PjMthKy(A, True))
End Function

Function MdMthNy(A As CodeModule, Optional MthNmPatn$ = ".", Optional IsNoMdNmPfx As Boolean, Optional Mdy0$) As String()
Dim Ay$(): Ay = SrcMthNy(MdSrc(A), MthNmPatn, Mdy0:=Mdy0)
If IsNoMdNmPfx Then
    MdMthNy = Ay
Else
    MdMthNy = AyAddPfx(Ay, MdNm(A) & ".")
End If
End Function

Function MdMthNyOfInproper(A As CodeModule, Optional ShwMsg As Boolean) As String()
If Not MdIsFunGpMd(A) Then
    If ShwMsg Then
        Debug.Print FmtQQ("MdMthNyOfInproper: Given Md should be begins with [M_] or [S_].  MdNm=[?]", MdNm(A))
    End If
    Exit Function ' M_Xxxx for Module with all pub-fun begins with Xxxx
End If                                             ' S_Xxxx for Module with single function of name=Xxxx
Dim Ay() As Mth
Dim Ay1() As Mth
    Ay = MdMthAy(A)
    Ay1 = AyWhPredNot(Ay, "MthIsInProperMd")
MdMthNyOfInproper = AyMapSy(Ay1, "MthNm")
End Function

Function MdNm$(A As CodeModule)
MdNm = A.Parent.Name
End Function

Function MdPj(A As CodeModule) As VBProject
Set MdPj = A.Parent.Collection.Parent
End Function

Function MdPjNm$(A As CodeModule)
MdPjNm = MdPj(A).Name
End Function

Function MdProperMdNy(A As CodeModule) As String()
Dim Ny$(): Ny = MdMthNy(A, IsNoMdNmPfx:=True, Mdy0:="Public")
MdProperMdNy = AyWhSingleEle(AyMapSy(Ny, "MthNm_ProperMdNm"))
End Function

Function MdRmk(A As CodeModule) As Boolean
Debug.Print "Rmk " & A.Parent.Name,
If MdIsAllRemarked(A) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To A.CountOfLines
    A.ReplaceLine J, "'" & A.Lines(J, 1)
Next
MdRmk = True
End Function

Function MdSrc(A As CodeModule) As String()
MdSrc = MdLy(A)
End Function

Function MdSrcExt$(A As CodeModule)
Dim O$
Select Case A.Parent.Type
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function

Function MdSrcFfn$(A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function

Function MdSrcFn$(A As CodeModule)
MdSrcFn = MdNm(A) & MdSrcExt(A)
End Function

Function MdSrtRpt(A As CodeModule) As DCRslt
Dim P$, M$
M = MdNm(A)
P = MdPjNm(A)
MdSrtRpt = SrcSrtRpt(MdSrc(A), P, M)
End Function

Function MdSrtRptLy(A As CodeModule) As String()
Dim P$: P = MdPjNm(A)
Dim M$: M = MdNm(A)
MdSrtRptLy = SrcSrtRptLy(MdSrc(A), P, M)
End Function

Function MdSrtedLines$(A As CodeModule)
MdSrtedLines = SrcSrtedLines(MdSrc(A))
End Function

Function MdTyNm$(A As CodeModule)
MdTyNm = CmpTy_Nm(MdCmpTy(A))
End Function

Function MdUnRmk(A As CodeModule) As Boolean
Debug.Print "UnRmk " & A.Parent.Name,
If Not MdIsAllRemarked(A) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
Next
MdUnRmk = True
End Function

Function MdXNm_Either(A) As Either
'Return ~.Left as MdDNm
'Or     ~.Right as PjNy() for those Pj holding giving Md
Dim P$, M$
Brk2Asg A, ".", P, M
If P <> "" Then
    MdXNm_Either = EitherL(A)
    Exit Function
End If
Dim Ny$()
Ny = VbeMdNmPjNy(CurVbe, M)
If Sz(Ny) = 1 Then
    MdXNm_Either = EitherL(Ny(0) & "." & M)
    Exit Function
End If
MdXNm_Either = EitherR(Ny)
End Function

Function Md_MthNm_z_ProperMdNm_S1S2Ay(A As CodeModule) As S1S2()
Dim Ny$(): Ny = MdMthNy(A, IsNoMdNmPfx:=True, Mdy0:="Public")
Md_MthNm_z_ProperMdNm_S1S2Ay = AyMapS1S2Ay(Ny, "MthNm_ProperMdNm")
End Function

Function Md_FunNy_OfPfx_ZZDash(A As CodeModule) As String()
Dim J%, O$(), L$, L1$, Is_ZFun As Boolean
For J = 1 To A.CountOfLines
    Is_ZFun = True
    L = A.Lines(J, 1)
    Select Case True
    Case IsPfx(L, "Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case IsPfx(L, "Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case Else:
        Is_ZFun = False
    End Select

    If Is_ZFun Then
        Push O, LinNm(L1)
    End If
Next
Md_FunNy_OfPfx_ZZDash = O
End Function

Function Md_Lines_ByFTNo$(A As CodeModule, X As FTNo)
Dim Cnt%: Cnt = FTNo_LinCnt(X)
If Cnt = 0 Then Exit Function
Md_Lines_ByFTNo = A.Lines(X.Fmno, Cnt)
End Function

Function Md_Ly_ByFTNo(A As CodeModule, X As FTNo) As String()
Md_Ly_ByFTNo = SplitCrLf(Md_Lines_ByFTNo(A, X))
End Function

Function Md_TstSub_BdyLines$(A As CodeModule)
Dim Ny$(): Ny = Md_FunNy_OfPfx_ZZDash(A)
If Sz(Ny) = 0 Then Exit Function
Ny = AySrt(Ny)
Dim O$()
Dim Pfx$
If A.Parent.Type = vbext_ct_ClassModule Then
    Pfx = "Friend "
End If
Push O, ""
Push O, Pfx & "Sub ZZZ__Tst()"
PushAy O, Ny
Push O, "End Sub"
Md_TstSub_BdyLines = Join(O, vbCrLf)
End Function

Function Md_TstSub_Lno%(A As CodeModule)
Dim J%
For J = 1 To A.CountOfLines
    If LinIsTstSub(A.Lines(J, 1)) Then Md_TstSub_Lno = J: Exit Function
Next
End Function

Function MdyIsSel(A$, MdySy$()) As Boolean
If Sz(MdySy) = 0 Then MdyIsSel = True: Exit Function
Dim Mdy
For Each Mdy In MdySy
    If Mdy = "Public" Then
        If A = "" Then MdyIsSel = True: Exit Function
    End If
    If A = Mdy Then MdyIsSel = True: Exit Function
Next
End Function

Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
O = A(0)
For J = 1 To UB(Av)
    If A(J) < O Then O = A(J)
Next
Min = O
End Function

Function MthDNm$(A As Mth)
MthDNm = MdDNm(A.Md) & "." & A.Nm
End Function

Function MthDNm_Lines$(A)
MthDNm_Lines = MthLines(MthDNm_Mth(A))
End Function

Function MthDNm_Mth(A) As Mth
Dim Ay$(): Ay = Split(A, ".")
Dim Nm$, M As CodeModule
Select Case Sz(Ay)
Case 1: Nm = Ay(0): Set M = CurMd
Case 2: Nm = Ay(1): Set M = Md(A)
Case 3: Nm = Ay(2): Set M = Md(Ay(0) & "." & Ay(1))
Case Else: Stop
End Select
Set MthDNm_Mth = Mth(M, Nm)
End Function

Function MthDNm_Nm$(A)
Dim Ay$(): Ay = Split(A, ".")
Dim Nm$
Select Case Sz(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthDNm_Nm = Nm
End Function

Function MthFNm$(A As Mth)
MthFNm = A.Nm & ":" & MdDNm(A.Md)
End Function

Function MthFNm_Mth(A) As Mth
Set MthFNm_Mth = MthDNm_Mth(MthFNm_MthDNm(A))
End Function

Function MthFNm_MthDNm$(A)
With Brk(A, ":")
    MthFNm_MthDNm = .S2 & "." & .S1
End With
End Function

Function MthFNm_Nm$(A$)
MthFNm_Nm = Brk(A, ":").S1
End Function

Function MthFmno%(A As Mth)
MthFmno = SrcMthFmix(MdSrc(A.Md), A.Nm) + 1
End Function

Function MthFmnoAy(A As Mth) As Integer()
MthFmnoAy = IntAy_Add1(SrcMthFmixAy(MdSrc(A.Md), A.Nm))
End Function

Function MthFTNoAy(A As Mth) As FTNo()
MthFTNoAy = SrcMthFTNoAy(MdSrc(A.Md), A.Nm)
End Function
Function MthFTNo(A As Mth) As FTNo
MthFTNo = SrcMthFTNo(MdSrc(A.Md), A.Nm)
End Function

Function MthIsExist(A As Mth) As Boolean
MthIsExist = MdMthFmno(A.Md, A.Nm) > 0
End Function

Function MthIsInProperMd(A As Mth) As Boolean
'Return True if mth is in a ProperMd
If Not MdIsFunGpMd(A.Md) Then MthIsInProperMd = True: Exit Function
Dim M$: M = MthNm_ProperMdNm(A.Nm): If M = "" Then MthIsInProperMd = True: Exit Function
MthIsInProperMd = M = MdNm(A.Md)
End Function

Function MthIsPub(A As Mth) As Boolean
Dim L$: L = MthLin(A)
If L = "" Then Stop
Dim Mdy$: Mdy = LinMdy(L)
If Mdy = "" Or Mdy = "Public" Then MthIsPub = True
End Function

Function MthKy_Sq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Sz(A) + 1, 1 To 6)
SqSetRow O, 1, FnyOf_MthKey
For J = 0 To UB(A)
    SqSetRow O, J + 2, Split(A(J), ":")
Next
MthKy_Sq = O
End Function

Function MthLCCOpt(A As Mth) As LCCOpt
Dim L%, C As LCCOpt
Dim M As CodeModule
Set M = A.Md
For L = M.CountOfDeclarationLines + 1 To M.CountOfLines
    Set C = LinLCCOpt(M.Lines(L, 1), A.Nm, L)
    If C.Som Then
        Set MthLCCOpt = LCCOpt(C.LCC)
        Exit Function
    End If
Next
End Function

Function MthLin$(A As Mth)
MthLin = SrcMthLin(MdSrc(A.Md), A.Nm)
End Function
Function MthBNm_MthNm$(A)
MthBNm_MthNm = TakBef(TakAftMust(A, "."), ":")
End Function
Function MthLinCnt%(A As Mth)
MthLinCnt = FTNoAy_LinCnt(MthFTNoAy(A))
End Function

Function MthLin_MthBrk(A) As MthBrk
Dim L$: L = A
Dim M$: M = LinShiftMdy(L)
Dim T$: T = LinShiftMthTy(L): If T = "" Then Stop
Dim N$: N = LinNm(L): If N = "" Then Stop
Set MthLin_MthBrk = MthBrk(N, M, T)
End Function

Function MthLin_MthNm$(A)
Dim L$: L = A
LinShiftMdy L
If LinShiftMthTy(L) = "" Then Exit Function
MthLin_MthNm = LinNm(L)
End Function

Function MthLin_MthANm$(A)
Dim L$, T$, T1$
L = A
LinShiftMdy L
T = LinShiftMthShtTy(L): If T = "" Then Exit Function
T1 = IIf(T = "Sub" Or T = "Fun", "", T)
MthLin_MthANm = LinNm(L) & IIf(T1 = "", "", ":") & T1
End Function

Function MthLin_MthKey$(A$, Optional PjNm$, Optional MdNm$, Optional IsWrap As Boolean)
Dim M$ 'Mdy
Dim T$ 'MthTy {Sub | Function | Function | Property Let | Property Set
Dim S$ 'MthShtTy *Sub *Fun *Get *Let *Set
Dim N$ 'Name
Dim IsMthLin As Boolean
    MthLin_BrkAsg A, IsMthLin, M, T, N
    If Not IsMthLin Then Stop
    S = MthTy_MthShtTy(T)
Dim P% 'Priority
    Select Case True
    Case IsPfx(N, "Init"): P = 1
    Case N = "ZZZ__Tst":    P = 9
    Case N = "ZZZZ__Tst":   P = 9
    Case IsPfx(N, "ZZZ_"): P = 9
    Case IsPfx(N, "ZZ_"):  P = 8
    Case IsPfx(N, "Z"):    P = 7
    Case Else:             P = 2
    End Select
Dim O$
    Dim Fmt$, NoPjNmMdNm As Boolean
    NoPjNmMdNm = PjNm = "" And MdNm = ""
    Fmt = IIf(NoPjNmMdNm, "?:?|?:?", "?:?|?:?|?:?")
    If Not IsWrap Then Fmt = Replace(Fmt, "|", ":")
    
    If NoPjNmMdNm Then
        O = FmtQQ(Fmt, P, N, S, M)
    Else
        O = FmtQQ(Fmt, PjNm, MdNm, P, N, S, M)
    End If

MthLin_MthKey = O
End Function

Function MthLines$(A As Mth)
MthLines = SrcMthLinesByNm(MdSrc(A.Md), A.Nm)
End Function

Function MthMdNm$(A As Mth)
MthMdNm = MdNm(A.Md)
End Function

Function MthNm$(A As Mth)
MthNm = A.Nm
End Function

Function MthNm_CmpLy(A, Optional InclSam As Boolean) As String()
Dim N$(): N = MthNm_DupFunFNy(A)
If Sz(N) > 1 Then
    MthNm_CmpLy = DupMthFNyGp_CmpLy(N, InclSam:=InclSam)
End If
End Function

Function MthNm_DupFunFNy(A) As String()
MthNm_DupFunFNy = VbeFunFNy(CurVbe, FunNmPatn:="^" & A & "$")
End Function

Function MthPjNm$(A As Mth)
MthPjNm = MdPjNm(A.Md)
End Function

Function MthProperMd(A As Mth) As CodeModule
'Mth here must be must belong to a StdMd
'Mth here must be Public, or,
'Mth name is ZZ_xxx, then it is ok to be private
If Not MdIsStdMd(A.Md) Then Stop
If Not IsPfx(A.Nm, "ZZ_") Then
    If Not MthIsPub(A) Then Stop
End If
Dim Pj As VBProject
Dim MdNm$
    MdNm = MthNm_ProperMdNm(A.Nm)
    Set Pj = MdPj(A.Md)
PjEnsMd Pj, MdNm
Set MthProperMd = PjMd(Pj, MdNm)
End Function

Function MthTy_IsVdt(A) As Boolean
MthTy_IsVdt = AyHas(SyOf_MthTy, A)
End Function

Function MthTy_MthShtTy$(A)
Dim O$
Select Case A
Case "Sub": O = A
Case "Function": O = "Fun"
Case "Property Get", "Property Let", "Property Set": O = LinRmvT1(A)
End Select
MthTy_MthShtTy = O
End Function
Function MdyShtMdy(A)
Dim O$
Select Case A
Case "", "Public":
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case Else: Stop
End Select
MdyShtMdy = O
End Function

Function NewA1() As Range
Set NewA1 = NewWs.Range("A1")
End Function

Function NewWb() As Workbook
Set NewWb = Xls.Workbooks.Add
End Function

Function NewWs() As Worksheet
Set NewWs = NewWb.Sheets(1)
End Function

Function OyPrpAy(Oy, PrpNm) As Variant()
OyPrpAy = OyPrpAyInto(Oy, PrpNm, EmpAy)
End Function

Function OyPrpAyInto(Oy, PrpNm, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(Oy) > 0 Then
    Dim I
    For Each I In Oy
        Push O, ObjPrp(I, PrpNm)
    Next
End If
OyPrpAyInto = O
End Function
Function ObjPrp(Obj, PrpNm)
On Error Resume Next
ObjPrp = CallByName(Obj, PrpNm, VbGet)
End Function

Function OyNy(Oy) As String()
Dim O$(): If Sz(Oy) = 0 Then Exit Function
Dim I
For Each I In Oy
    Push O, CallByName(I, "Name", VbGet)
Next
OyNy = O
End Function

Function OyToStrSy(A) As String()
If Sz(A) = 0 Then Exit Function
Dim O$()
ReDim O(UB(A))
Dim J&
For J = 0 To UB(A)
    O(J) = A(J).ToStr
Next
OyToStrSy = O
End Function
Private Function ZZZ_LinShiftXXX()
Dim O$: O = "AA{|}BB "
Debug.Assert LinShiftXXX(O, "{|}") = "AA"
Debug.Assert O = "BB "
End Function
Function LinShiftXXX$(O$, XXX$)
Dim P%: P = InStr(O, XXX)
If P = 0 Then Exit Function
LinShiftXXX = Left(O, P - 1)
O = Mid(O, P + Len(XXX))
End Function
Function LinShiftDTerm$(O$)
LinShiftDTerm = LinShiftXXX(O, ".")
End Function
Function AWhPred(A As Mth, PredFunNm$)
Dim O: O = A: Erase O
Dim I
If Sz(A) > 0 Then
    For Each I In A
        If Run(PredFunNm, I) Then
            PushObj O, I
        End If
    Next
End If
AWhPred = O
End Function
Function DicAddKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
If A.Count > 0 Then
    For Each K In A.Keys
        O.Add Pfx & K, A(K)
    Next
End If
Set DicAddKeyPfx = O
End Function
Function Pj(PjNm$) As VBProject
Set Pj = CurVbe.VBProjects(PjNm)
End Function

Function PjClsAy(A As VBProject, Optional ClsNmPatn$ = ".", Optional ExclClsNy0) As CodeModule()
PjClsAy = PjMbrAy(A, ClsNmPatn, ExclClsNy0, Array(vbext_ct_ClassModule))
End Function

Function PjClsMdAy(A As VBProject, Optional MbrNmPatn$ = ".", Optional ExclMbrNy0) As CodeModule()
PjClsMdAy = PjMbrAy(A, MbrNmPatn, ExclMbrNy0, Array(vbext_ct_ClassModule, vbext_ct_StdModule))
End Function

Function PjClsMdNy(A As VBProject, Optional Patn$ = ".", Optional ExclNy0) As String()
PjClsMdNy = PjMbrNy(A, Patn, ExclNy0, Array(vbext_ct_ClassModule, vbext_ct_StdModule))
End Function

Function PjClsNy(A As VBProject, Optional Patn$ = ".", Optional ExclNy0) As String()
PjClsNy = PjMbrNy(A, Patn, ExclNy0, Array(vbext_ct_ClassModule))
End Function

Function PjCmp(A As VBProject, Nm) As VBComponent
Set PjCmp = A.VBComponents(CStr(Nm))
End Function

Function PjDicOfMthKeyzzzMthLines(A As VBProject) As Dictionary
Dim I
Dim O As New Dictionary
For Each I In PjMbrAy(A)
    Set O = DicAdd(O, MdDicOfMthKeyzzzMthLines(CvMd(I)))
Next
Set PjDicOfMthKeyzzzMthLines = O
End Function

Function PjDupFunFNy(A As VBProject, Optional IsSamMthBdyOnly As Boolean) As String()
Dim N$(): N = PjFunFNy(A)
Dim N1$(): N1 = FunFNy_DupFunFNy(N)
If IsSamMthBdyOnly Then
    N1 = DupFunFNy_SamMthBdyFunFNy(N1, A)
End If
PjDupFunFNy = N1
End Function

Function PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.Filename
End Function

Function PjFstMd(A As VBProject) As CodeModule
Dim Cmp As VBComponent, O$()
For Each Cmp In A.VBComponents
    If Cmp.Type = vbext_ct_StdModule Then
        Set PjFstMd = Cmp.CodeModule
        Exit Function
    End If
Next
For Each Cmp In A.VBComponents
    If Cmp.Type = vbext_ct_ClassModule Then
        Set PjFstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function

Function PjFunBdyDic(A As VBProject) As Dictionary
Stop '
End Function

Function PjFunFNy(A As VBProject, Optional MdNmPatn$ = ".", Optional FunNmPatn$ = ".", Optional ExclFunNy0$, Optional Mdy0$) As String()
Dim Ay() As CodeModule
    Ay = PjMdAy(A, MdNmPatn:=MdNmPatn) ' Note: Fun is exist Md, so PjMdAy is used
If Sz(Ay) = 0 Then Exit Function
Dim O$(), I
For Each I In Ay
    PushAy O, MdFunFNy(CvMd(I), FunNmPatn:=FunNmPatn, ExclFunNy0:=ExclFunNy0, Mdy0:=Mdy0)
Next
PjFunFNy = O
End Function

Function PjFunNy(A As VBProject, Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".") As String()
Dim Ay() As CodeModule: Ay = PjMbrAy(A, MbrNmPatn)
If Sz(Ay) = 0 Then Exit Function
Dim I, O$()
For Each I In Ay
    PushAy O, MdMthNy(CvMd(I), MthNmPatn)
Next
O = AyAddPfx(O, A.Name & ".")
PjFunNy = O
End Function

Function PjFunPfxAy(A As VBProject) As String()
Dim Ay() As CodeModule: Ay = PjMdAy(A)
Dim Ay1(): Ay1 = AyMap(Ay, "MdFunPfxAy")
PjFunPfxAy = AyOfAy_Ay(Ay1)
End Function

Function PjHasCmp(A As VBProject, Nm$) As Boolean
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If Cmp.Name = Nm Then PjHasCmp = True: Exit Function
Next
End Function
Sub LocStr_Go(A)
LocGo LocStr_Loc(A)
End Sub
Function LocStr_Loc(A) As Loc

End Function
Sub LocGo(A As Loc)

End Sub
Function ItrMapInto(A, MapFunNm$, OIntoAy)
Dim I, O
O = OIntoAy: Erase O
For Each I In A
    Push O, Run(MapFunNm, I)
Next
ItrMapInto = O
End Function

Function ItrMap(A, MapFunNm$)
ItrMap = ItrMapInto(A, MapFunNm, EmpAy)
End Function

Function ItrMapSy(A, MapFunNm$) As String()
ItrMapSy = ItrMapInto(A, MapFunNm, EmpSy)
End Function

Function PjRfNy(A As VBProject) As String()
PjRfNy = ItrNy(A.References)
End Function
Function PjHasRfNm(A As VBProject, RfNm$) As Boolean
Dim I, R As Reference
For Each I In A.References
    Set R = I
    If R.Name = RfNm Then PjHasRfNm = True: Exit Function
Next
End Function
Function PjHasRfFfn(A As VBProject, RfFfn) As Boolean
Dim I, R As Reference
For Each I In A.References
    Set R = I
    If R.FullPath = RfFfn Then PjHasRfFfn = True: Exit Function
Next
End Function
Sub ZZ_ItrWhPrpItr()
Dim Act1, Act2, Act3, CmpItr, MbrItr
Set MbrItr = Pj("QTool").VBComponents
Set Act1 = ItrWhPrpItr(MbrItr, "Type", ApItr(vbext_ct_StdModule))
Set Act2 = ItrWhPrpItr(MbrItr, "Type", ApItr(vbext_ct_ClassModule))
Set Act3 = ItrWhPrpItr(MbrItr, "Type", ApItr(vbext_ct_ClassModule, vbext_ct_StdModule))
Stop
End Sub
Function ItrWhPrpItr(A, PrpNm$, WhItr As Collection) As Collection
If IsNothing(WhItr) Then Set ItrWhPrpItr = A: Exit Function
Dim I, O As New Collection, P
For Each I In A
    For Each P In WhItr
        If ObjPrp(I, PrpNm) = P Then
            O.Add I
        End If
    Next
Next
Set ItrWhPrpItr = O
End Function
Function PjMbrItr(A As VBProject, Optional MbrNmPatn$ = ".", Optional ExclMbrLikNy As Collection, Optional CmpTyItr As Collection) As Collection
Dim Itr As Collection: Set Itr = ItrWhNmPatnExcl(A.VBComponents, MbrNmPatn, ExclMbrLikNy)
Set PjMbrItr = ItrWhPrpItr(Itr, "Type", CmpTyItr)
End Function

Function PjMbrAy(A As VBProject, Optional MbrNmPatn$ = ".", Optional ExclMbrNy0, Optional CmpTyAy) As CodeModule()
Dim MdAy() As CodeModule
PjMbrAy = AyMapPXInto(PjMbrNy(A, MbrNmPatn, ExclMbrNy0, CmpTyAy), "PjMd", A, MdAy)
End Function
Function ItrWhNmPatn(A, NmPatn$) As Collection
'Assume A is collection of object-with-name-property
If NmPatn = "." Then Set ItrWhNmPatn = ItrClone(A): Exit Function
Dim R As RegExp: Set R = Re(NmPatn)
Dim I, O As New Collection
For Each I In A
    If R.Test(I) Then
        O.Add I
    End If
Next
Set ItrWhNmPatn = O
End Function
Function ItrClone(A) As Collection
Dim I, O As New Collection
For Each I In A
    O.Add I
Next
Set ItrClone = O
End Function
Function ItrWhExclLikNmItr(A, ExclLikNmItr As Collection) As Collection
'Assume A is Object Itr with property-name
If IsNothing(ExclLikNmItr) Then Set ItrWhExclLikNmItr = ItrClone(A): Exit Function
Dim I, O As New Collection, Nm
For Each I In A
    For Each Nm In ExclLikNmItr
        If Not I.Name Like Nm Then
            O.Add I
        End If
    Next
Next
Set ItrWhExclLikNmItr = O
End Function
Function ItrWhNmPatnExcl(A, Optional NmPatn$ = ".", Optional ExclLikNmItr As Collection) As Collection
'Assume A is collection of object with Name-property else break
Set ItrWhNmPatnExcl = ItrWhExclLikNmItr(ItrWhNmPatn(A, NmPatn), ExclLikNmItr)
End Function

Function PjMbrNy(A As VBProject, Optional Patn$ = ".", Optional ExclNy0, Optional CmpTyAy) As String()
Dim Ny$(): Ny = ItrNy(A.VBComponents, Patn, ExclNy0)
If IsMissing(CmpTyAy) Then
    PjMbrNy = Ny
    Exit Function
End If
Dim O$(), N
For Each N In Ny
    If AyHas(CmpTyAy, PjCmp(A, N).Type) Then
        Push O, N
    End If
Next
PjMbrNy = O
End Function

Function PjMd(A As VBProject, Nm) As CodeModule
Set PjMd = PjCmp(A, Nm).CodeModule
End Function
Function PjMdItr(A As VBProject, Optional MdNmPatn$ = ".", Optional ExclMdLikNmItr As Collection) As Collection
Set PjMdItr = PjMbrItr(A, MdNmPatn, ExclMdLikNmItr, ApItr(vbext_ct_StdModule))
End Function

Function PjMdAy(A As VBProject, Optional MdNmPatn$ = ".", Optional ExclMdNy0) As CodeModule()
PjMdAy = PjMbrAy(A, MdNmPatn, ExclMdNy0, Array(vbext_ct_StdModule))
End Function

Function PjMdNy(A As VBProject, Optional Patn$ = ".", Optional ExclNy0) As String()
PjMdNy = PjMbrNy(A, Patn, ExclNy0, Array(vbext_ct_StdModule))
End Function

Function PjMdNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_StdModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
PjMdNy_With_TstSub = O
End Function

Function PjMdSrtRpt(A As VBProject) As MdSrtRpt
'SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
Dim Ay() As CodeModule: Ay = PjMbrAy(A)
Dim Ny$(): Ny = OyNy(Ay)
Dim LyAy()
Dim IsSam() As Boolean
    Dim J%, R As DCRslt
    For J = 0 To UB(Ay)
        R = MdSrtRpt(Ay(J))
        Push LyAy, DCRsltLy(R)
        Push IsSam, DCRsltIsSam(R)
    Next
With PjMdSrtRpt
    Set .RptDic = AyPair_Dic(Ny, LyAy)
    .MdNy = PjMdSrtRpt_1(Ny, IsSam)
End With
End Function

Function PjMdSrtRpt_1(MdNy$(), IsSam() As Boolean) As String()
Dim O$(), J%
For J = 0 To UB(MdNy)
    Push O, AlignL(MdNy(J), 30) & " " & IsSam(J)
Next
PjMdSrtRpt_1 = O
End Function

Function PjMd_and_Cls_Ny(A As VBProject) As String()
Dim O$(), Cmp As VBComponent
For Each Cmp In A.VBComponents
    If Cmp.Type = vbext_ct_StdModule Or Cmp.Type = vbext_ct_ClassModule Then
        Push O, Cmp.Name
    End If
Next
PjMd_and_Cls_Ny = O
End Function

Function PjMthAy(A As VBProject, Optional MdNmPatn$ = ".", Optional MthNmPatn$ = ".", Optional Mdy0$) As Mth()
Dim Ay() As CodeModule: Ay = PjMdAy(A, MdNmPatn)
Dim M, O() As Mth
For Each M In Ay
    PushObjAy O, MdMthAy(CvMd(M), MdNmPatn, Mdy0)
Next
PjMthAy = O
End Function

Function PjMthKy(A As VBProject, Optional IsWrap As Boolean) As String()
PjMthKy = AyMapPXSy(PjMbrAy(A), "MdMthKy", IsWrap)
End Function

Function PjMthKySq(A As VBProject) As Variant()
PjMthKySq = MthKy_Sq(PjMthKy(A, True))
End Function

Function PjMthNy(A As VBProject, Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".", Optional Mdy0$) As String()
Dim Ay() As CodeModule: Ay = PjClsMdAy(A, MbrNmPatn)
If Sz(Ay) = 0 Then Exit Function
Dim I, O$()
For Each I In Ay
    PushAy O, MdMthNy(CvMd(I), MthNmPatn, Mdy0:=Mdy0)
Next
O = AyAddPfx(O, A.Name & ".")
PjMthNy = O
End Function

Function PjMthNyOfInproper(A As VBProject) As String()
Dim I, O$()
Dim Ay() As CodeModule: Ay = PjMdAy(A)
If Sz(Ay) = 0 Then Exit Function
Dim N$, M As CodeModule
For Each I In Ay
    Set M = CvMd(I)
    PushAy O, AyAddPfx(MdMthNyOfInproper(M), MdDNm(M) & ".")
Next
PjMthNyOfInproper = O
End Function

Function PjPth$(A As VBProject)
PjPth = FfnPth(A.Filename)
End Function

Function PjRfAy(A As VBProject) As Reference()
PjRfAy = ItrAy(A.References, EmpRfAy)
End Function

Function PjRfCfgFfn(A As VBProject)
PjRfCfgFfn = PjSrcPth(A) & "PjRf.Cfg"
End Function

Function PjRfLy(A As VBProject) As String()
Dim RfAy() As Reference
    RfAy = PjRfAy(A)
Dim O$()
Dim Ny$(): Ny = OyNy(RfAy)
Ny = AyAlignL(Ny)
Dim J%
For J = 0 To UB(Ny)
    Push O, Ny(J) & " " & RfFfn(RfAy(J))
Next
PjRfLy = O
End Function

Function PjSrcPth(A As VBProject)
Dim Ffn$: Ffn = PjFfn(A)
If Ffn = "" Then Exit Function
Dim Fn$: Fn = FfnFn(Ffn)
Dim P$: P = FfnPth(A.Filename)
If P = "" Then Exit Function
Dim O$:
O = P & "Src\": PthEns O
O = O & Fn & "\":                  PthEns O
PjSrcPth = O
End Function

Function PjSrtRptLy(A As VBProject) As String()
Dim Ay() As CodeModule: Ay = PjMbrAy(A)
Dim O$(), I, M As CodeModule
For Each I In Ay
    Set M = I
    PushAy O, MdSrtRptLy(M)
Next
PjSrtRptLy = O
End Function

Function PjSrtRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
Dim A1 As MdSrtRpt
A1 = PjMdSrtRpt(A)
Dim O As Workbook: Set O = DicWb(A1.RptDic)
Dim Ws As Worksheet
Set Ws = WbAddWs(O, "Md Idx")
'Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
'LoCol_LnkWs Lo, "Md"
'If Vis Then WbVis O
'Set PjSrtRptWb = O
Stop '
End Function

Function Pj_ClsNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_ClassModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
Pj_ClsNy_With_TstSub = O
End Function

Function Pj_TstClass_Bdy$(A As VBProject)
Dim N1$() ' All Class Ny with 'Friend Sub ZZZ__Tst' method
Dim N2$()
Dim A1$, A2$
Const Q1$ = "Sub ?()|Dim A As New ?: A.ZZZ__Tst|End Sub"
Const Q2$ = "Sub ?()|#.?.ZZZ__Tst|End Sub"
N1 = Pj_ClsNy_With_TstSub(A)
A1 = SeedExpand(Q1, N1)
N2 = PjMdNy_With_TstSub(A)
A2 = Replace(SeedExpand(Q2, N2), "#", A.Name)
Pj_TstClass_Bdy = A1 & vbCrLf & A2
End Function

'Function PthEntAy(A, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute, Optional IsRecursive As Boolean) As String()
'If Not IsRecursive Then
'    PthEntAy = AyAdd(PthSubPthAy(A), PthFfnAy(A, FilSpec, Atr))
'    Exit Function
'End If
'Erase O
'PthPushEntAyR A
'PthEntAy = O
'Erase O
'End Function
'Function PthFdr$(A$)
'Ass PthHasPthSfx(A)
'Dim P$: P = RmvLasChr(A)
'PthFdr = TakAftRev(A, "\")
'End Function
Function PthFfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthFfnAy = AyAddPfx(PthFnAy(A, Spec, Atr), A)
End Function

Function PthFfnItr(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As Collection
Set PthFfnItr = CollAddPfx(PthFnItr(A, Spec, Atr), A)
End Function

Function PthFnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
If Not PthIsExist(A) Then
    Debug.Print FmtQQ("PthFnAy: Given Path(?) does not exit", A)
    Exit Function
End If
Dim O$()
Dim M$
M = Dir(A & Spec)
If Atr = 0 Then
    While M <> ""
       Push O, M
       M = Dir
    Wend
    PthFnAy = O
End If
Ass PthHasPthSfx(A)
While M <> ""
    If GetAttr(A & M) And Atr Then
        Push O, M
    End If
    M = Dir
Wend
PthFnAy = O
End Function

Function PthFnItr(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As Collection
Set PthFnItr = AyItr(PthFnAy(A, Spec, Atr))
End Function

Function PthHasPthSfx(A) As Boolean
PthHasPthSfx = LasChr(A) = "\"
End Function

Function PthIsExist(A) As Boolean
Ass PthHasPthSfx(A)
PthIsExist = Fso.FolderExists(A)
End Function

Function Re(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set Re = O
End Function

Function RfFfn$(A As Reference)
On Error Resume Next
RfFfn = A.FullPath
End Function

Function PjRfNm_RfFfn$(A As VBProject, RfNm$)
PjRfNm_RfFfn = PjPth(A) & RfNm & ".xlam"
End Function

Function RgLo(A As Range, Optional LoNm$) As ListObject
Dim O As ListObject
Set O = RgWs(A).ListObjects.Add(xlSrcRange, A, , xlYes)
If LoNm <> "" Then O.Name = LoNm
Set RgLo = O
End Function

Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgWs(A As Range)
Set RgWs = A.Parent
End Function

Function RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Function

Function RmvLasChr$(A)
RmvLasChr = Left(A, Len(A) - 1)
End Function

Function RmvLasNChr$(A, N%)
RmvLasNChr = Left(A, Len(A) - N)
End Function

Function RmvPfx$(A, Pfx)
If IsPfx(A, Pfx) Then
    RmvPfx = Mid(A, Len(Pfx) + 1)
Else
    RmvPfx = A
End If
End Function

Function RplDblSpc$(A)
Dim O$: O = Trim(A)
Dim J&
While HasSubStr(O, "  ")
    J = J + 1: If J > 10000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function RplPun$(A)
Dim O$(), J&, L&, C$
L = Len(A)
If L = 0 Then Exit Function
ReDim O(L - 1)
For J = 1 To L
    C = Mid(A, J, 1)
    If IsPun(C) Then
        O(J - 1) = " "
    Else
        O(J - 1) = C
    End If
Next
RplPun = Join(O, "")
End Function

Function RplVBar$(A)
RplVBar = Replace(A, "|", vbCrLf)
End Function

Function S1S2Ay_Add(A() As S1S2, B() As S1S2) As S1S2()
Dim O() As S1S2
Dim J&
O = A
For J = 0 To UB(B)
    PushObj O, B(J)
Next
S1S2Ay_Add = O
End Function

Function S1S2Ay_Dic(A() As S1S2) As Dictionary
Dim J&, O As New Dictionary
For J = 0 To UB(A)
    O.Add A(J).S1, A(J).S2
Next
Set S1S2Ay_Dic = O
End Function

Function S1S2Ay_FmtLy(A() As S1S2) As String()
Dim W1%: W1 = S1S2Ay_S1LinesWdt(A)
Dim W2%: W2 = S1S2Ay_S2LinesWdt(A)
Dim W%(1)
W(0) = W1
W(1) = W2
Dim H$: H = WdtAy_HdrLin(W)
S1S2Ay_FmtLy = S1S2Ay_LinesLinesLy(A, H, W1, W2)
End Function

Function S1S2Ay_LinesLinesLy(A() As S1S2, H$, W1%, W2%) As String()
Dim O$(), I&
Push O, H
For I = 0 To UB(A)
   PushAy O, S1S2_Ly(A(I), W1, W2)
   Push O, H
Next
S1S2Ay_LinesLinesLy = O
End Function

Function S1S2Ay_S1LinesWdt%(A() As S1S2)
S1S2Ay_S1LinesWdt = LinesAy_Wdt(S1S2Ay_Sy1(A))
End Function

Function S1S2Ay_S2LinesWdt%(A() As S1S2)
S1S2Ay_S2LinesWdt = LinesAy_Wdt(S1S2Ay_Sy2(A))
End Function

Function S1S2Ay_Sy1(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S1
Next
S1S2Ay_Sy1 = O
End Function

Function S1S2Ay_Sy2(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S2
Next
S1S2Ay_Sy2 = O
End Function

Function S1S2_Ly(A As S1S2, W1%, W2%) As String()
Dim S1$(), S2$()
S1 = SplitCrLf(A.S1)
S2 = SplitCrLf(A.S2)
Dim M%, J%, O$(), Lin$, A1$, A2$, U1%, U2%
    U1 = UB(S1)
    U2 = UB(S2)
    M = Max(U1, U2)
Dim Spc1$, Spc2$
    Spc1 = Space(W1)
    Spc2 = Space(W2)
For J = 0 To M
   If J > U1 Then
       A1 = Spc1
   Else
       A1 = StrAlignL(S1(J), W1)
   End If
   If J > U2 Then
       A2 = Spc2
   Else
       A2 = StrAlignL(S2(J), W2)
   End If
   Lin = "| " + A1 + " | " + A2 + " |"
   Push O, Lin
Next
S1S2_Ly = O
End Function

Function SeedExpand$(QVbl$, Ny$())
Dim O$()
Dim Sy$(): Sy = SplitVBar(QVbl)
Dim J%, I
For J = 0 To UB(Ny)
    For Each I In Sy
       Push O, Replace(I, "?", Ny(J))
    Next
Next
SeedExpand = JnCrLf(O)
End Function

Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function

Function SplitSsl(A) As String()
SplitSsl = Split(RplDblSpc(Trim(A)), " ")
End Function

Function SplitVBar(Vbl$) As String()
SplitVBar = Split(Vbl, "|")
End Function

Function SqWs(A, Optional Vis As Boolean) As Worksheet
Dim A1 As Range: Set A1 = NewA1
CellPutSq A1, A
RgVis A1, Vis
Set SqWs = RgWs(A1)
End Function

Function SrcMthBrkAy(A$()) As MthBrk()
Dim L$(): L = SrcMthLinAy(A)
Dim X() As MthBrk
SrcMthBrkAy = AyMapInto(L, "MthLin_MthBrk", X)
End Function

Function SrcAllMthFmnoAy(A$()) As Integer()
Dim N%(): N = SrcAllMthFmixAy(A)
Dim J%
For J = 0 To UB(N)
    N(J) = N(J) + 1
Next
SrcAllMthFmnoAy = N
End Function

Function SrcAllMthFmixAy(A$()) As Integer()
Dim J%, O%()
For J = 0 To UB(A)
    If LinIsMthLin(A(J)) Then
        Push O, J
    End If
Next
SrcAllMthFmixAy = O
End Function

Function SrcAllMthFTIxAy(A$()) As FTIx()
Dim F%(): F = SrcAllMthFmixAy(A$)
Dim N%: N = Sz(F)
If N = 0 Then Exit Function
Dim O() As FTIx
ReDim O(N - 1)
Dim J%
For J = 0 To N - 1
    Set O(J) = FTIx(F(J), SrcMthToix(A, F(J)))
Next
SrcAllMthFTIxAy = O
End Function

Function SrcMthLinAy(A$()) As String()
Dim L%(): L = SrcAllMthFmixAy(A)
If Sz(L) = 0 Then Exit Function
Dim O$(), LL
For Each LL In L
    Push O, SrcContLin(A, CInt(LL))
Next
SrcMthLinAy = O
End Function

Function SrcAllMthNy(A$()) As String()
Dim L, O$(), Nm$
For Each L In A
    Nm = LinMthNm(L)
    If Nm <> "" Then
        PushNoDup O, Nm
    End If
Next
SrcAllMthNy = O
End Function

Function SrcContLin$(A$(), Lx%)
Dim O$(), J%, L$
For J = Lx To UB(A)
    L = A(J)
    If Right(L, 2) <> " _" Then
        Push O, L
        SrcContLin = Join(O, "")
        Exit Function
    End If
    Push O, RmvLasNChr(L, 2)
Next
ErImposs
End Function

Function SrcDclLinCnt%(A$())
Dim I&
    I = SrcFstMthLx(A)
    If I = -1 Then
        SrcDclLinCnt = Sz(A)
        Exit Function
    End If
    I = SrcMthRmkLx(A, I)
Dim O&, L$
    For I = I - 1 To 0 Step -1
        If LinIsCd(A(I)) Then
            O = I + 1
            GoTo X
        End If
    Next
X:
SrcDclLinCnt = O
End Function

Function SrcDclLines$(A$())
SrcDclLines = Join(SrcDclLy(A), vbCrLf)
End Function

Function SrcDclLy(A$()) As String()
If Sz(A) = 0 Then Exit Function
Dim N&
   N = SrcDclLinCnt(A)
If N <= 0 Then Exit Function
SrcDclLy = AyFstNEle(A, N)
End Function

Function SrcDicOfMthKeyzzzMthLines(A$(), Optional PjNm$, Optional MdNm$, Optional ExclDcl As Boolean) As Dictionary
Dim L%(): L = SrcAllMthFmixAy(A)
Dim K$
Dim O As New Dictionary
    If Not ExclDcl Then
        If PjNm = "" And MdNm = "" Then
            K$ = PjNm & "." & MdNm & ".*Dcl"
        Else
            K = "*Dcl"
        End If
        O.Add K, SrcDclLines(A)
    End If
    If Sz(L) = 0 Then GoTo X
    Dim MthNm$, Lin$, Lines$, Lx
    For Each Lx In L
        Lin = SrcContLin(A, CInt(Lx))
        MthNm = LinMthNm(Lin):               If MthNm = "" Then Stop
        Lines = SrcMthLinesByMthFmix(A, Lx): If Lines = "" Then Stop
        K = MthLin_MthKey(Lin, PjNm, MdNm)
        O.Add K, Lines
    Next
X:
Set SrcDicOfMthKeyzzzMthLines = O
End Function

'abc
'xyz
Function SrcDicOfMthNmzzzMthLines(A$(), Optional ExclDcl As Boolean) As Dictionary
Dim L%(): L = SrcAllMthFmixAy(A)
Dim O As New Dictionary
    If Not ExclDcl Then O.Add "*Dcl", SrcDclLines(A)
    If Sz(L) = 0 Then GoTo X
    Dim MthNm$, Lin$, Lines$, Lx
    For Each Lx In L
        Lin = A(Lx)
        MthNm = LinMthNm(Lin):            If MthNm = "" Then Stop
        Lines = SrcMthLinesByMthFmix(A, Lx): If Lines = "" Then Stop
        If O.Exists(MthNm) Then
            If LinPrpSubFun(Lin) <> "Property" Then Stop
            O(MthNm) = O(MthNm) & vbCrLf & vbCrLf & Lines
        Else
            O.Add MthNm, Lines
        End If
    Next
X:
Set SrcDicOfMthNmzzzMthLines = O
End Function

Function SrcEndLx(A$(), MthLx)
Dim F$: F = "End " & LinMthTy(A(MthLx))
Dim J%
For J = MthLx + 1 To UB(A)
    If IsPfx(A(J), F) Then SrcEndLx = J: Exit Function
Next
Stop
End Function

Function SrcFstMthLx&(A$())
Dim J%
For J = 0 To UB(A)
   If LinIsMthLin(A(J)) Then
       SrcFstMthLx = J
       Exit Function
   End If
Next
SrcFstMthLx = -1
End Function

Function SrcMthFmnoAy(A$(), MthNm) As Integer()
Dim O%(): O = SrcMthFmixAy(A, MthNm)
Dim J%
For J = 0 To UB(O)
    O(J) = O(J) + 1
Next
SrcMthFmnoAy = O
End Function

Function SrcMthFmix%(A$(), MthNm, Optional Fmix% = 0)
Dim J%, L$
For J = Fmix To UB(A)
    L = SrcContLin(A, J)
    If LinMthNm(L) = MthNm Then
        SrcMthFmix = J
        Exit Function
    End If
Next
SrcMthFmix = -1
End Function

Function SrcMthFTIx(A$(), MthNm) As FTIx
Dim Fmix%: Fmix = SrcMthFmix(A, MthNm)
Dim Toix%: Toix = SrcMthToix(A, Fmix)
Set SrcMthFTIx = FTIx(Fmix, Toix)
End Function

Function SrcMthFmixAy(A$(), MthNm) As Integer()
Dim L%
L = SrcMthFmix(A, MthNm): If L <= 0 Then Exit Function
Dim O%(): Push O, L
Dim S$: S = A(L)
If LinPrpSubFun(S) = "Property" Then
    L = SrcMthFmix(A, MthNm, L + 1)
    If L > 0 Then Push O, L
End If
SrcMthFmixAy = O
End Function

Function SrcMthFTNoAy(A$(), MthNm) As FTNo()
Dim X() As FTNo
Dim Ay() As FTIx: Ay = SrcMthFTIxAy(A, MthNm)
SrcMthFTNoAy = AyMapInto(Ay, "FTIx_FTNo", X)
End Function

Function SrcMthFTNo(A$(), MthNm) As FTNo
SrcMthFTNo = FTIx_FTNo(SrcMthFTIx(A, MthNm))
End Function

Function SrcMthFTIxAy(A$(), MthNm) As FTIx()
Dim F%()
F = SrcMthFmixAy(A, MthNm): If Sz(F) <= 0 Then Exit Function
Dim O() As FTIx
ReDim O(UB(F))
Dim J%
For J = 0 To UB(F)
    Set O(J) = FTIx(F(J), SrcMthToix(A, F(J)))
Next
SrcMthFTIxAy = O
End Function

Function SrcMthLin$(A$(), MthNm)
Dim L%: L = SrcMthFmix(A, MthNm)
SrcMthLin = SrcContLin(A, L)
End Function

Function SrcMthLinesByMthFmix$(A$(), MthFmix)
Dim P1$
    P1 = SrcMthRmkLines(A, MthFmix)
Dim P2$
    Dim L2%
    L2 = SrcEndLx(A, MthFmix): If L2 = 0 Then Stop
    P2 = Join(AyWhFmTo(A, MthFmix, L2), vbCrLf)
If P1 = "" Then
    SrcMthLinesByMthFmix = P2
Else
    SrcMthLinesByMthFmix = P1 & vbCrLf & P2
End If
End Function

Function SrcMthLinesByNm$(A$(), MthNm)
Dim L%(): L = SrcMthFmixAy(A, MthNm)
If Sz(L) = 0 Then Exit Function
Dim MthLx, O$()
For Each MthLx In L
    Push O, SrcMthLinesByMthFmix(A, MthLx)
Next
SrcMthLinesByNm = Join(O, vbCrLf & vbCrLf)
End Function

Function SrcMthNy(A$(), Optional MthNmPatn$ = ".", Optional ExclMthNy0$, Optional Mdy0$) As String()
Dim L%(): L = SrcAllMthFmixAy(A)
If Sz(L) = 0 Then Exit Function
Dim ExclMthNy$(): ExclMthNy = DftNy(ExclMthNy0)
Dim O$()
    Dim MdySy$(): MdySy = DftMdySy(Mdy0)
    Dim MthLx, Lin$, N$, R As RegExp, M$
    Set R = Re(MthNmPatn)
    For Each MthLx In L
        Lin = A(MthLx)
        N = LinMthNm(Lin)
        If AyHas(ExclMthNy, N) Then GoTo Nxt
        If R.Test(N) Then
            M = LinMdy(Lin)
            If MdyIsSel(M, MdySy) Then
                PushNoDup O, N
            End If
        End If
Nxt:
    Next
SrcMthNy = O
End Function

Function SrcMthRmkLines$(A$(), MthLx)
Dim O$(), J%, L$, I%
Dim Lx&: Lx = SrcMthRmkLx(A, MthLx)

For J = Lx To MthLx - 1
    L = Trim(A(J))
    If L = "" Or L = "'" Then GoTo X
    If Left(L, 1) <> "'" Then Stop
    Push O, L
X:
Next
SrcMthRmkLines = Join(O, vbCrLf)
End Function

Function SrcMthRmkLx&(A$(), MthLx)
Dim M1&
    Dim J&
    For J = MthLx - 1 To 0 Step -1
        If LinIsCd(A(J)) Then
            M1 = J
            GoTo M1IsFnd
        End If
    Next
    M1 = -1
M1IsFnd:
Dim M2&
    For J = M1 + 1 To MthLx - 1
        If Trim(A(J)) <> "" Then
            M2 = J
            GoTo M2IsFnd
        End If
    Next
    M2 = MthLx
M2IsFnd:
SrcMthRmkLx = M2
End Function

Function SrcMthToix%(A$(), Fmix%)
Dim T$: T = LinPrpSubFun(A(Fmix))
If Not AyHas(SyOf_PrpSubFun, T) Then Stop
Dim B$: B = "End " & T
Dim J%
For J = Fmix + 1 To UB(A)
    If IsPfx(A(J), B) Then
        SrcMthToix = J
        Exit Function
    End If
Next
Stop
End Function

Function SrcSrtRpt(A$(), PjNm$, MdNm$) As DCRslt
Dim B$(): B = SrcSrtedLy(A)
Dim A1 As Dictionary
Dim B1 As Dictionary
Set A1 = SrcDicOfMthKeyzzzMthLines(A, PjNm, MdNm)
Set B1 = SrcDicOfMthKeyzzzMthLines(B, PjNm, MdNm)
Dim O As DCRslt: O = DicCmp(A1, B1, "BefSrt", "AftSrt")
SrcSrtRpt = O
End Function

Function SrcSrtRptLy(A$(), PjNm$, MdNm$) As String()
SrcSrtRptLy = DCRsltLy(SrcSrtRpt(A, PjNm, MdNm))
End Function

Function SrcSrtedBdyLines$(A$())
If Sz(A) = 0 Then Exit Function
Dim D As Dictionary
Dim D1 As Dictionary
    Set D = SrcDicOfMthKeyzzzMthLines(A, ExclDcl:=True)
    Set D1 = DicSrt(D)
Dim O$()
    Dim K
   For Each K In D1.Keys
       Push O, vbCrLf & D1(K)
   Next
SrcSrtedBdyLines = JnCrLf(O)
End Function

Function SrcSrtedBdyLy(A$())
SrcSrtedBdyLy = SplitCrLf(SrcSrtedBdyLines(A))
End Function

Function SrcSrtedLines$(A$())
SrcSrtedLines = JnCrLf(SrcSrtedLy(A))
End Function

Function SrcSrtedLy(A$()) As String()
Dim A1$(), A2$()
A1 = SrcDclLy(A)
A2 = SrcSrtedBdyLy(A)
SrcSrtedLy = AyAddAp(A1, A2)
End Function

Function SslSy(Ssl) As String()
SslSy = Split(Trim(RplDblSpc(Ssl)), " ")
End Function
Function StrAlignL$(S$, W, Optional ErIFmnotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "StrAlignL"
Dim L%: L = Len(S)
If L > W Then
    If ErIFmnotEnoughWdt Then
        Stop
        'Er CSub, "Len({S)) > {W}", S, W
    End If
    If DoNotCut Then
        StrAlignL = S
        Exit Function
    End If
End If

If W >= L Then
    StrAlignL = S & Space(W - L)
    Exit Function
End If
If W > 2 Then
    StrAlignL = Left(S, W - 2) + ".."
    Exit Function
End If
StrAlignL = Left(S, W)
End Function

Function StrDup$(S, N%)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function StrLin$(A)
StrLin = A
End Function

Function StrNy(A) As String()
Dim O$: O = RplPun(A)
Dim O1$(): O1 = AyWhSingleEle(SslSy(O))
Dim O2$()
Dim J%
For J = 0 To UB(O1)
    If Not IsDigit(FstChr(O1(J))) Then Push O2, O1(J)
Next
StrNy = O2
End Function

Function StrSubStrCnt&(A$, SubStr$)
Dim P&, O%, L%
L = Len(SubStr)
P = 1
While P > 0
    P = InStr(P, A, SubStr)
    If P > 0 Then O = O + 1: P = P + L
Wend
StrSubStrCnt = O
End Function

Function SyOf_Mdy() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(0) = "Private"
    O(1) = "Friend"
    O(2) = "Public"
End If
SyOf_Mdy = O
End Function

Function SyOf_MthTy() As String()
Static O$(4), A As Boolean
If Not A Then
    A = True
    O(0) = "Property Get"
    O(1) = "Property Let"
    O(2) = "Property Set"
    O(3) = "Sub"
    O(4) = "Function"
End If
SyOf_MthTy = O
End Function
Function SyOf_MthShtTy() As String()
Static O$(4), A As Boolean
If Not A Then
    A = True
    O(0) = "Get"
    O(1) = "Let"
    O(2) = "Set"
    O(3) = "Sub"
    O(4) = "Fun"
End If
SyOf_MthShtTy = O
End Function

Function SyOf_PrpSubFun() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(0) = "Property"
    O(1) = "Sub"
    O(2) = "Function"
End If
SyOf_PrpSubFun = O
End Function

Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function TmpFcsv$(Optional Fdr$, Optional Fnn$)
TmpFcsv = TmpFfn(".csv", Fdr, Fnn)
End Function

Function TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
If Fnn0 = "" Then
    Fnn = TmpNm
Else
    Fnn = Fnn0
End If
TmpFfn = TmpPth(Fdr) & Fnn & Ext
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpPth$(Optional Fdr$)
Dim X$
   If Fdr <> "" Then
       X = Fdr & "\"
   End If
Dim O$
   O = TmpPthHom & X:   PthEns O
   O = O & TmpNm & "\": PthEns O
   PthEns O
TmpPth = O
End Function

Function TmpPthHom$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthHom = X
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function
Function IsLinesAy(A) As Boolean
If Not IsSy(A) Then Exit Function
If Sz(A) = 0 Then Exit Function
Dim S
For Each S In A
    If IsLines(S) Then IsLinesAy = True: Exit Function
Next
End Function
Function IsLines(A) As Boolean
IsLines = True
If HasSubStr(A, vbCr) Then Exit Function
If HasSubStr(A, vbLf) Then Exit Function
IsLines = False
End Function
Function LinesAy_Lines$(A)
Stop
Dim W%
W = LinesAy_Wdt(A): If W = 0 Then Exit Function
LinesAy_Lines = Join(A, Space(W))
End Function
Function ObjToStr$(A)
If Not IsObject(A) Then Stop
On Error GoTo X
ObjToStr = A.ToStr: Exit Function
X: ObjToStr = QuoteSqBkt(TypeName(A))
End Function
Function QuoteSqBkt$(A)
QuoteSqBkt = "[" & A & "]"
End Function
Function LvlSep$(Lvl%)
Select Case Lvl
Case 0: LvlSep = "."
Case 1: LvlSep = "-"
Case 2: LvlSep = "+"
Case 3: LvlSep = "="
Case 4: LvlSep = "*"
Case Else: LvlSep = Lvl
End Select
End Function
Sub ZZ_VarStr()
Dim A: A = Array(SslSy("sdf sdf df"), SslSy("sdf sdf"))
Debug.Print VarStr(A)
End Sub
Function VarStr$(A, Optional Lvl%)
Dim T$, S$, W%, I, O$(), Sep
Select Case True
Case IsPrim(A): VarStr = A
Case IsLinesAy(A): VarStr = LinesAy_Lines(A)
Case IsSy(A): VarStr = JnCrLf(A)
Case IsNothing(A): VarStr = "#Nothing"
Case IsEmpty(A): VarStr = "#Empty"
Case IsMissing(A): VarStr = "#Missing"
Case IsObject(A)
    VarStr = ObjToStr(A)
    T = TypeName(A)
    Select Case T
    Case "CodeModule"
        Dim M As CodeModule
        Set M = A
        VarStr = FmtQQ("*Md{?}", M.Parent.Name)
        Exit Function
    End Select
    VarStr = "*" & T
    Exit Function
Case IsArray(A)
    If Sz(A) = 0 Then Exit Function
    For Each I In A
        Push O, VarStr(I, Lvl + 1)
    Next
    W = LinesAy_Wdt(O)
    Sep = LvlSep(Lvl)
    VarStr = Join(O, vbCrLf & StrDup(Sep, W) & vbCrLf)
Case Else
End Select
End Function

Function VbeAllPjNy(A As Vbe) As String()
VbeAllPjNy = ItrNy(A.VBProjects)
End Function

Function VbeDupFunCmpLy(A As Vbe, Optional InclSam As Boolean) As String()
Dim N$(): N = VbeDupFunFNy(A)
Dim Ay(): Ay = DupFunFNy_GpAy(N)
Dim O$(), J%
Push O, FmtQQ("Total ? dup function.  ? of them has mth-lines are same", Sz(Ay), DupFunFNyGpAy_AllSameCnt(Ay))
Dim Cnt%, Sam%
For J = 0 To UB(Ay)
    PushAy O, DupMthFNyGp_CmpLy(Ay(J), Cnt, Sam, InclSam:=InclSam)
Next
VbeDupFunCmpLy = O
End Function

Function VbeDupFunDrs(A As Vbe, Optional IsNoSrt As Boolean, Optional PjNmPatn$ = ".", Optional ExclPjNy0, Optional IsSamMthBdyOnly As Boolean) As Drs
Dim Fny$(), Dry()
Fny = SplitSsl("Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src")
Dry = VbeDupFunDry(A, PjNmPatn:=PjNmPatn, ExclPjNy0:=ExclPjNy0, IsSamMthBdyOnly:=IsSamMthBdyOnly)
Set VbeDupFunDrs = Drs(Fny, Dry)
End Function

Function VbeDupFunDry(A As Vbe, Optional PjNmPatn$, Optional ExclPjNy0, Optional IsSamMthBdyOnly As Boolean) As Variant()
Dim N$(): N = VbeFunFNy(A, PjNmPatn:=PjNmPatn, ExclPjNy0:=ExclPjNy0)
Dim N1$(): N1 = FunFNy_DupFunFNy(N)
    If IsSamMthBdyOnly Then
        N1 = DupFunFNy_SamMthBdyFunFNy(N1, A)
    End If
Dim GpAy()
    GpAy = DupFunFNy_GpAy(N1)
    If Sz(GpAy) = 0 Then Exit Function
Dim O()
    Dim Gp
    For Each Gp In GpAy
        PushAy O, DupFunFNyGp_Dry(CvSy(Gp))
    Next
VbeDupFunDry = O
End Function

Function VbeDupFunFNy(A As Vbe, Optional IsNoSrt As Boolean, Optional ExclPjNy0, Optional IsSamMthBdyOnly As Boolean) As String()
Dim N$(): N = VbeFunFNy(A, ExclPjNy0:=ExclPjNy0, ExclFunNy0:="ZZZ__Tst")
Dim N1$(): N1 = FunFNy_DupFunFNy(N)
If IsSamMthBdyOnly Then
    N1 = DupFunFNy_SamMthBdyFunFNy(N1, A)
End If
VbeDupFunFNy = N1
End Function

Function VbeDupMdNy(A As Vbe) As String()
Dim O$()
Stop '
VbeDupMdNy = O
End Function

Function VbeFstQPj(A As Vbe) As VBProject
Dim I
For Each I In A.VBProjects
    If FstChr(CvPj(I).Name) = "Q" Then
        Set VbeFstQPj = I
        Exit Function
    End If
Next
End Function

Function VbeFunFNy(A As Vbe, Optional PjNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional FunNmPatn$ = ".", Optional ExclPjNy0, Optional ExclMdNy0, Optional ExclFunNy0$, Optional Mdy0$) As String()
Dim Ay() As VBProject
    Ay = VbePjAy(A, PjNmPatn, ExclPjNy0)
If Sz(Ay) = 0 Then Exit Function
Dim O$(), I
For Each I In Ay
    PushAy O, PjFunFNy(CvPj(I), MdNmPatn:=MdNmPatn, FunNmPatn:=FunNmPatn, ExclFunNy0:=ExclFunNy0, Mdy0:=Mdy0$)
Next
VbeFunFNy = O
End Function

Function VbeFunPfxAy(A As Vbe) As String()
Dim O$(), P
For Each P In VbePjAy(A)
    PushAyNoDup O, PjFunPfxAy(CvPj(P))
Next
VbeFunPfxAy = O
End Function

Function VbeMdNmPjNy(A As Vbe, MdNm$) As String()
Dim I, O$()
For Each I In VbePjAy(A)
    If PjHasCmp(CvPj(I), MdNm) Then
        Push O, CvPj(I).Name
    End If
Next
VbeMdNmPjNy = O
End Function

Function VbeMthKy(A As Vbe, Optional IsWrap As Boolean) As String()
Dim O$(), I
For Each I In VbePjAy(A)
    PushAy O, PjMthKy(CvPj(I), IsWrap)
Next
VbeMthKy = O
End Function

Function VbeMthNy(A As Vbe, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$) As String()
Dim Ay() As VBProject: Ay = VbePjAy(A, MdNmPatn)
If Sz(Ay) = 0 Then Exit Function
Dim I, O$()
For Each I In Ay
    PushAy O, PjMthNy(CvPj(I), MthNmPatn, MdNmPatn, Mdy)
Next
VbeMthNy = O
End Function

Function VbeMthNyOfInproper(A As Vbe) As String()
Dim I, O$()
For Each I In VbePjAy(A)
    PushAy O, PjMthNyOfInproper(CvPj(I))
Next
VbeMthNyOfInproper = O
End Function

Function VbePjAy(A As Vbe, Optional PjNmPatn$ = ".", Optional ExclPjNy0) As VBProject()
Dim N$(): N = VbePjNy(A, PjNmPatn, ExclPjNy0)
Dim PjAy() As VBProject
VbePjAy = AyMapInto(N, "Pj", PjAy)
End Function

Function VbePjNy(A As Vbe, Optional Patn$ = ".", Optional ExclNy0) As String()
VbePjNy = AyWhPatn(VbeAllPjNy(A), Patn, ExclNy0)
End Function

Function VbeSrcPth(A As Vbe)
Dim Pj As VBProject:
Set Pj = VbeFstQPj(A)
Dim Ffn$: Ffn = PjFfn(Pj)
If Ffn = "" Then Exit Function
VbeSrcPth = FfnPth(Pj.Filename)
End Function

Function VbeSrtRptLy(A As Vbe) As String()
Dim Ay() As VBProject: Ay = VbePjAy(A)
Dim O$(), I, M As VBProject
For Each I In Ay
    Set M = I
    PushAy O, PjSrtRptLy(M)
Next
VbeSrtRptLy = O
End Function

Function VblLines$(A)
VblLines = Replace(A, "|", vbCrLf)
End Function

Function WbAddWs(A As Workbook, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet
Set O = A.Sheets.Add(A.Sheets(1))
If WsNm <> "" Then
   O.Name = WsNm
End If
Set WbAddWs = O
End Function

Function WbCn_TxtCn(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set WbCn_TxtCn = A.TextConnection
End Function

Function WbTxtCn(A As Workbook) As TextConnection
Dim N%: N = WbTxtCnCnt(A)
If N <> 1 Then
    Stop
    Exit Function
End If
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then
        Set WbTxtCn = C.TextConnection
        Exit Function
    End If
Next
ErImposs
End Function

Function WbTxtCnCnt%(A As Workbook)
Dim C As WorkbookConnection, Cnt%
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then Cnt = Cnt + 1
Next
WbTxtCnCnt = Cnt
End Function

Function WbTxtCnStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
If IsNothing(T) Then Exit Function
WbTxtCnStr = T.Connection
End Function

Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function

Function WdtAy_HdrLin$(A%())
Dim O$(), W
For Each W In A
    Push O, StrDup("-", W + 2)
Next
WdtAy_HdrLin = "|" + Join(O, "|") + "|"
End Function

Function WinOf_Imm() As VBIDE.Window
Set WinOf_Imm = WinTy_Win(vbext_wt_Immediate)
End Function

Function WinOf_Lcl() As VBIDE.Window
Set WinOf_Lcl = WinTy_Win(vbext_wt_Locals)
End Function

Function WinTy_Win(Ty As vbext_WindowType) As VBIDE.Window
Set WinTy_Win = CurVbe.Windows(Ty)
End Function
Sub RgBdrTop(A As Range)
RgBdr A, xlEdgeTop
End Sub

Sub RgBdr(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub
Function RgR(A As Range, R) As Range
Set RgR = RgRCRC(A, R, 1, R, RgNCol(A))
End Function
Function RgC(A As Range, C) As Range
Set RgC = RgRCRC(A, 1, RgNRow(A), 1, C)
End Function
Function RgNRow&(A As Range)
RgNRow = A.Rows.Count
End Function
Function RgNCol%(A As Range)
RgNCol = A.Columns.Count
End Function
Sub RgBdrAround(A As Range)
A.BorderAround XlLineStyle.xlContinuous, xlMedium
If A.Row > 1 Then RgBdrBottom RgR(A, 0)
If A.Column > 1 Then RgBdrRight RgC(A, 0)
RgBdrTop RgR(A, RgNRow(A) + 1)
RgBdrLeft RgC(A, RgNCol(A) + 1)
End Sub

Sub RgBdrBottom(A As Range)
RgBdr A, xlEdgeBottom
End Sub

Sub RgBdrInside(A As Range)
RgBdr A, xlInsideHorizontal
RgBdr A, xlInsideVertical
End Sub

Sub RgBdrLeft(A As Range)
RgBdr A, xlEdgeLeft
If A.Column > 1 Then
    RgBdr RgC(A, 0), xlEdgeRight
End If
End Sub

Sub RgBdrRight(A As Range)
RgBdr A, xlEdgeRight
If A.Column < MaxCol Then
    RgBdr RgC(A, A.Column + 1), xlEdgeLeft
End If
End Sub

Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Cells(1, 1)
End Function

Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function

Function WsRR(A As Worksheet, R1, R2) As Range
Set WsRR = A.Range(WsRC(A, R1, 1), WsRC(A, R2, 1)).EntireRow
End Function

Function WsVis(A As Worksheet) As Worksheet
XlsVis A.Application
Set WsVis = A
End Function

Function Xls() As Excel.Application
Static Y As Excel.Application
On Error GoTo X
Dim A$: A = Y.Name
Set Xls = Y
Exit Function
X:
Set Y = New Excel.Application
Set Xls = Y
End Function

Function XlsHasAddInFn(A As Excel.Application, AddInFn) As Boolean
Dim I As Excel.AddIn
Dim N$: N = UCase(AddInFn)
For Each I In A.AddIns
    If UCase(I.Name) = N Then XlsHasAddInFn = True: Exit Function
Next
End Function

Sub Asg(V, OV)
If IsObject(V) Then
   Set OV = V
Else
   OV = V
End If
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub

Sub AyBrw(Ay, Optional Fnn$)
Dim T$
T = TmpFt("AyBrw", Fnn)
AyWrt Ay, T
FtBrw T
End Sub

Sub AyDmp(A)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Debug.Print I
Next
End Sub
Sub AyDoPX(A, DoMthNm$, P)
If Sz(A) = 0 Then Exit Sub
Dim X
For Each X In A
    Run DoMthNm, P, X
Next
End Sub
Sub AyDoXP(A, DoMthNm$, P)
If Sz(A) = 0 Then Exit Sub
Dim X
For Each X In A
    Run DoMthNm, X, P
Next
End Sub

Sub AyDo(A, DoMthNm$)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Run DoMthNm, I
Next
End Sub

Sub AyWrt(A, Ft$)
StrWrt JnCrLf(A), Ft
End Sub

Sub Brk2Asg(A, Sep$, O1$, O2$)
Dim P%: P = InStr(A, Sep)
If P = 0 Then
    O1 = ""
    O2 = Trim(A)
Else
    O1 = Trim(Left(A, P - 1))
    O2 = Trim(Mid(A, P + 1))
End If
End Sub

Sub BrkAsg(A, Sep$, O1, O2)
With Brk(A, Sep)
    O1 = .S1
    O2 = .S2
End With
End Sub

Sub CmpRmv(A As VBComponent)
A.Collection.Remove A
End Sub

Sub DDNmBrkAsg(A, O1$, O2$, O3$)
Dim Ay$(): Ay = Split(A, ".")
Select Case Sz(Ay)
Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
Case Else: Stop
End Select
End Sub
Function TyNm$(A)
TyNm = TypeName(A)
End Function
Sub DicTyBrw(A As Dictionary)
DicBrw DicTy(A)
End Sub
Function DicTy(A As Dictionary) As Dictionary
Set DicTy = DicMap(A, "TyNm")
End Function
Sub DicBrw(A As Dictionary)
WsVis S1S2Itr_Ws(DicS1S2Itr(A))
End Sub

Sub DrsBrw(A As Drs)
Stop '
End Sub

Sub DupFunFNy_ShwNotDupMsg(A$(), MthNm)
Select Case Sz(A)
Case 0: Debug.Print FmtQQ("DupFunFNy_ShwNotDupMsg: no such Fun(?) in CurVbe", MthNm)
Case 1
    Dim B As S1S2: Set B = Brk(A(0), ":")
    If B.S1 <> MthNm Then Stop
    Debug.Print FmtQQ("DupFunFNy_ShwNotDupMsg: Fun(?) in Md(?) does not have dup-Fun", B.S1, B.S2)
End Select
End Sub

Sub ErImposs()
Stop ' Impossible
End Sub

'Function DftFfn(Ffn0, Optional Ext$ = ".txt", Optional Pth0$, Optional Fdr$)
'If Ffn0 <> "" Then DftFfn = Ffn0: Exit Function
'Dim Pth$: Pth = DftPth(Pth0)
'DftFfn = Pth & TmpNm & Ext
'End Function
'Function DftPth$(Optional Pth0$, Optional Fdr$)
'If Pth0 <> "" Then DftPth = Pth0: Exit Function
'DftPth = TmpPth(Fdr)
'End Function
'Function FfnAddFnSfx(A$, Sfx$)
'FfnAddFnSfx = FfnRmvExt(A) & Sfx & FfnExt(A)
'End Function
Sub FfnCpyToPth(A, ToPth$, Optional OvrWrt As Boolean)
Fso.CopyFile A, ToPth$ & FfnFn(A), OvrWrt
End Sub

Sub FfnDlt(A)
On Error GoTo X
Kill A
Exit Sub
X: Debug.Print FmtQQ("FfnDtl: Kill(?) Er(?)", A, Err.Description)
End Sub

Sub FtBrw(A)
Shell "code.cmd """ & A & """", vbHide
'Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub

Sub FtRmvFst4Lines(Ft$)
Dim A$: A = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
Dim B$: B = Left(A, 55)
Dim C$: C = Mid(A, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If B <> B1 Then Stop
Fso.CreateTextFile(Ft, True).Write C
End Sub

Sub FunFNm_BrkAsg(A$, OFunNm$, OPjNm$, OMdNm$)
With Brk(A, ":")
    OFunNm = .S1
    With Brk(.S2, ".")
        OPjNm = .S1
        OMdNm = .S2
    End With
End With
End Sub

Sub FunNm_Cmp(A, Optional InclSam As Boolean)
AyDmp FunNm_CmpLy(A, InclSam)
End Sub

Sub FunSync(A As Mth, Optional ShwCmpLyAft As Boolean)
Dim Lines$: Lines = MthLines(A)
If Lines = "" Then
    Debug.Print FmtQQ("Give Mth(?) not exist", MthDNm(A))
    Exit Sub
End If
Dim M() As Mth
    M = FunSync__1(A, Lines) ' Mth to be replaced
If Sz(M) = 0 Then Exit Sub
Dim I
For Each I In M
    MthRpl CvMth(I), Lines
Next
If ShwCmpLyAft Then
    FunNm_Cmp A.Nm, True
End If
End Sub

Sub FxaNm_Crt(A)
FxaCrt FxaNm_Fxa(A)
End Sub

Sub FxaCrt(A)
If FfnIsExist(A) Then
    Debug.Print FmtQQ("FxaCrt: Fxa(?) is already exist", A)
    Exit Sub
End If
If XlsHasAddInFn(CurXls, FfnFn(A)) Then Stop: Exit Sub
Dim O As Workbook
Set O = CurXls.Workbooks.Add
O.SaveAs A, XlFileFormat.xlOpenXMLAddIn
O.Close
Dim AddIn As AddIn: Set AddIn = CurXls.AddIns.Add(A)
AddIn.Installed = True
Dim Pj As VBProject
Set Pj = VbePjFfn_Pj(CurVbe, A)
Pj.Name = FfnFnn(A)
PjSav Pj
End Sub
Function VbePjFfn_Pj(A As Vbe, Ffn) As VBProject
Dim I
For Each I In A.VBProjects ' Cannot use VbePjAy(A), should use A.VBProjects
                           ' due to VbePjAy(X).FileName gives error
                           ' but (Pj in A.VBProjects).FileName is OK
    Debug.Print PjFfn(CvPj(I))
    If StrIsEq(PjFfn(CvPj(I)), Ffn) Then
        Set VbePjFfn_Pj = I
        Exit Function
    End If
Next
End Function
Function XlsAddIn(A As Excel.Application, FxaNm) As Excel.AddIn
Dim I As Excel.AddIn
For Each I In A.AddIns
    If StrIsEq(I.Name, FxaNm & ".xlam") Then Set XlsAddIn = I
Next
End Function
Function StrIsEq(A, B) As Boolean
StrIsEq = StrComp(A, B, vbTextCompare) = 0
End Function
Sub ItrDoSub(A, SubNm$)
Dim I
For Each I In A
    CallByName A, SubNm, VbMethod
Next
End Sub

Sub MdAddFun(A As CodeModule, Nm$, IsFun As Boolean)
Dim L$
    Dim B$
    B = IIf(IsFun, "Function", "Sub")
    L = FmtQQ("? ?()|End ?", B, Nm, B)
MdAppLines A, L
MthGo Mth(A, Nm)
End Sub

Sub MdAppLines(A As CodeModule, Lines$)
A.InsertLines A.CountOfLines + 1, Lines
End Sub

Sub MdClr(A As CodeModule, Optional IsSilent As Boolean)
With A
    If .CountOfLines = 0 Then Exit Sub
    If Not IsSilent Then Debug.Print FmtQQ("MdClr: Md(?) of lines(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub MdClsWin(A As CodeModule)
A.CodePane.Window.Close
End Sub

Sub MdCmp(A As CodeModule, B As CodeModule)
Dim A1 As Dictionary, B1 As Dictionary
    Set A1 = MdDic(A)
    Set B1 = MdDic(B)
Dim C As DCRslt
    C = DicCmp(A1, B1, MdDNm(A), MdDNm(B))
AyBrw DCRsltLy(C)
End Sub

Sub MdCmpByNm(MdDNm1$, MdDNm2$)
MdCmp Md(MdDNm1), Md(MdDNm2)
End Sub

Sub MdCpy(A As CodeModule, ToPj As VBProject)
Dim MdNm$
Dim FmPj As VBProject
    Set FmPj = MdPj(A)
    MdNm = A.Parent.Name
If PjHasCmp(ToPj, MdNm) Then
    Debug.Print FmtQQ("MdCpy: Md(?) exists in TarPj(?).  Skip moving", MdNm, ToPj.Name)
    Exit Sub
End If
Dim TmpFil$
    TmpFil = TmpFfn(".txt")
    Dim SrcCmp As VBComponent
    Set SrcCmp = A.Parent
    SrcCmp.Export TmpFil
    If SrcCmp.Type = vbext_ct_ClassModule Then
        FtRmvFst4Lines TmpFil
    End If
Dim TarCmp As VBComponent
    Set TarCmp = ToPj.VBComponents.Add(A.Parent.Type)
    TarCmp.CodeModule.AddFromFile TmpFil
Kill TmpFil
PjSav ToPj
Debug.Print FmtQQ("MdCpy: Md(?) is moved from SrcPj(?) to TarPj(?).", MdNm, FmPj.Name, ToPj.Name)
End Sub

Sub MdDlt(A As CodeModule)
Dim M$, P$, Pj As VBProject
    M = MdNm(A)
    Set Pj = MdPj(A)
    P = Pj.Name
A.Parent.Collection.Remove A.Parent
PjSav Pj
Debug.Print FmtQQ("MdDlt: Md(?) is deleted from Pj(?)", M, P)
End Sub

Sub MdEndTrim(A As CodeModule, Optional ShwMsg As Boolean)
If A.CountOfLines = 0 Then Exit Sub
Dim N$: N = MdDNm(A)
Dim J%
While Trim(A.Lines(A.CountOfLines, 1)) = ""
    If ShwMsg Then Debug.Print FmtQQ("MdEndTrim: Lin(?) in Md(?) is removed due to it is blank", A.CountOfLines, N)
    A.DeleteLines A.CountOfLines, 1
    If A.CountOfLines = 0 Then Exit Sub
    If J > 1000 Then Stop
    J = J + 1
Wend
End Sub

Sub MdExport(A As CodeModule)
Dim F$: F = MdSrcFfn(A)
A.Parent.Export F
Debug.Print MdNm(A)
End Sub

Sub MdGo(A As CodeModule)
Cls_Win_ExcptImm
With A.CodePane
    .Show
    .Window.WindowState = vbext_ws_Maximize
End With
SendKeys "%WV"
End Sub

Sub MdGoLCCOpt(Md As CodeModule, LCCOpt As LCCOpt)
MdGo Md
With LCCOpt
    If .Som Then
        With .LCC
            Md.CodePane.TopLine = .Lno
            Md.CodePane.SetSelection .Lno, .C1, .Lno, .C2
        End With
    End If
End With
SendKeys "^{F4}"
End Sub

Sub MdRmvFTNo(A As CodeModule, X As FTNo)
A.DeleteLines X.Fmno, FTNo_LinCnt(X)
End Sub

Sub MdRmvFTNoAy(A As CodeModule, X() As FTNo)
Dim J%
For J = UB(X) To 0 Step -1
    MdRmvFTNo A, X(J)
Next
End Sub

Sub MdRplCxt(A As CodeModule, Cxt$)
Dim N%: N = A.CountOfLines
MdClr A, IsSilent:=True
A.AddFromString Cxt
Debug.Print FmtQQ("MdRpl_Cxt: Md(?) of Ty(?) of Old-LinCxt(?) is replaced by New-Len(?) New-LinCnt(?).<-----------------", _
    MdDNm(A), MdTyNm(A), N, Len(Cxt), LinesLinCnt(Cxt))
End Sub

Sub MdSrt(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
Debug.Print "Sorting: "; AlignL(Nm, 30); " ";
If MdNm(A) = "G_Tool" And MdPjNm(A) = "QTool" Then
    Debug.Print "<<<< Skipped"
    Exit Sub
End If
Dim NewLines$: NewLines = MdSrtedLines(A)
Dim Old$: Old = MdLines(A)
'Exit if same
    If Old = NewLines Then
        Debug.Print "<== Same"
        Exit Sub
    End If
Debug.Print "<-- Sorted";
'Delete
    Debug.Print FmtQQ("<--- Deleted (?) lines", A.CountOfLines);
    MdClr A, IsSilent:=True
'Add sorted lines
    A.AddFromString NewLines
    Md_Rmv_EmptyLines_AtEnd A
    Debug.Print "<----Sorted Lines added...."
End Sub

Sub Md_FunNm_z_ProperMdNm_Brw(A As CodeModule)
S1S2Ay_Brw Md_MthNm_z_ProperMdNm_S1S2Ay(A)
End Sub

Sub Md_Gen_TstSub(A As CodeModule)
Md_Rmv_TstSub A
MdAppLines A, Md_TstSub_BdyLines(A)
End Sub

Sub Md_Mov_ToPj(A As CodeModule, ToPj As VBProject)
If MdNm(A) = "F__Tool" And CurPj.Name = "QTool" Then
    Debug.Print "Md(QTool.F__Tool) cannot be moved"
    Exit Sub
End If
MdCpy A, ToPj
MdDlt A
End Sub

Sub Md_Rmv_EmptyLines_AtEnd(A As CodeModule)
Dim J%
While A.CountOfLines > 1
    J = J + 1
    If J > 10000 Then Stop
    If Trim(A.Lines(A.CountOfLines, 1)) <> "" Then Exit Sub
    A.DeleteLines A.CountOfLines, 1
Wend
End Sub

Sub Md_Rmv_TstSub(A As CodeModule)
Dim L&, N&
L = Md_TstSub_Lno(A)
If L = 0 Then Exit Sub
Dim Fnd As Boolean, J%
For J = L + 1 To A.CountOfLines
    If IsPfx(A.Lines(J, 1), "End Sub") Then
        N = J - L + 1
        Fnd = True
        Exit For
    End If
Next
If Not Fnd Then Stop
A.DeleteLines L, N
End Sub

Sub MthBrkAsg(A As Mth, OMdy$, OMthTy$)
Dim L$: L = MthLin(A)
OMdy = LinMdy(L)
OMthTy = LinMthTy(L)
End Sub

Sub MthCpy(A As Mth, ToMd As CodeModule, Optional IsSilent As Boolean)
If MdHasMth(ToMd, A.Nm) Then
    Debug.Print FmtQQ("MthCpy_ToMd: Fm-Mth(?) is Found in To-Md(?)", A.Nm, MdNm(ToMd))
    Exit Sub
End If
If ObjPtr(A.Md) = ObjPtr(ToMd) Then
    Debug.Print FmtQQ("MthCpy_ToMd: Fm-Mth-Md(?) cannot be To-Md(?)", MthMdNm(A), MdNm(ToMd))
    Exit Sub
End If
MdAppLines ToMd, MthLines(A)
If Not IsSilent Then
    Debug.Print FmtQQ("MthCpy_ToMd: Mth(?) is copied ToMd(?)", MthDNm(A), MdDNm(ToMd))
End If
End Sub

Sub MthCpyToPj(A As Mth, ToPj As VBProject, Optional IsSilent As Boolean)
Dim ToMdNm$: ToMdNm = MthNm_ProperMdNm(A.Nm)
Dim ToMd As CodeModule: Set ToMd = PjMd(ToPj, ToMdNm)
MthCpy A, ToMd
End Sub

Sub MthDNm_Mov_ToProperMd(A)
MthMovToProperMd MthDNm_Mth(A)
End Sub
Function VbeMthIdAy(A As Vbe) As String()
Dim Ay(): Ay = AyMap(VbePjAy(A), "PjMthIdAy")
VbeMthIdAy = AyOfAy_Ay(Ay)
End Function
Sub ZZ_PjMthIdAy()
AyBrw PjMthIdAy(CurPj)
End Sub
Sub ZZ_VbeMthIdAy()
AyBrw VbeMthIdAy(CurVbe)
End Sub
Function PjMthIdAy(A As VBProject) As String()
Dim Ay(): Ay = AyMap(PjMbrAy(A), "MdMthIdAy")
PjMthIdAy = AyOfAy_Ay(Ay)
End Function
Sub MthGo(A As Mth)
MdGoLCCOpt A.Md, MthLCCOpt(A)
End Sub

Sub ZZ_MdMthIdAy()
AyBrw MdMthIdAy(CurMd)
End Sub

Function MdMthIdAy(A As CodeModule, Optional ExclMdy As Boolean) As String()
Dim L$(): L = SrcMthLinIdAy(MdSrc(A), InclMdy:=Not ExclMdy): If Sz(L) = 0 Then Exit Function
MdMthIdAy = AyAddPfx(L, MdDNm(A) & ".")
End Function
Sub ZZ_SrcMthLinIdAy()
AyBrw SrcMthLinIdAy(CurSrc, InclMdy:=True)
End Sub
Function SrcMthLinIdAy(A$(), InclMdy As Boolean) As String()
Dim L$(): L = SrcMthLinAy(A): If Sz(L) = 0 Then Exit Function
SrcMthLinIdAy = AyMapXPSy(L, "MthLin_MthLinId", InclMdy)
End Function
Function MthLin_MthLinId$(A, InclMdy As Boolean)
'MthLinId : MthNm:ShtMdy
Dim L$: L = A
Dim M$: M = LinShiftShtMdy(L)
Dim T$: T = LinShiftMthShtTy(L): If T = "" Then Exit Function
Dim N$: N = LinNm(L)
If InclMdy Then
    MthLin_MthLinId = N & ":" & T & ":" & M
Else
    MthLin_MthLinId = N & ":" & T
End If
End Function
Sub MthLin_BrkAsg(A$, Optional OIsMthLin As Boolean, Optional OMdy$, Optional OMthTy$, Optional OMthNm$)
OIsMthLin = False
Dim L$: L = A
OMdy = LinShiftMdy(L)
OMthTy = LinShiftMthTy(L): If Not AyHas(SyOf_MthTy, OMthTy) Then Stop
OMthNm = LinNm(L)
OIsMthLin = True
End Sub

Sub MthMov(A As Mth, ToMd As CodeModule)
MthCpy A, ToMd, IsSilent:=True
MthRmv A, IsSilent:=True
Debug.Print FmtQQ("MthMov: Mth(?) is moved to Md(?)", MthDNm(A), MdDNm(ToMd))
MdClsWin ToMd
End Sub

Sub MthMovToProperMd(A As Mth)
If MdCmpTy(A.Md) <> vbext_ct_StdModule Then
    Debug.Print FmtQQ("MthMovToProperMd: Md(?) in not in StdMd", MthDNm(A))
    Exit Sub
End If
If Not IsPfx(A.Nm, "ZZ_") Then
    If Not MthIsPub(A) Then
        Debug.Print FmtQQ("MdMovToProperMd: Mth(?) is not public", MthDNm(A))
        Exit Sub
    End If
End If
MthMov A, MthProperMd(A)
End Sub

Sub MthNm_Cmp(A)
Debug.Print "MthNm: Don't use MthNm_Cmp, but use FunNm_Cmp"
AyDmp MthNm_CmpLy(A)
End Sub

Sub MthRmv(A As Mth, Optional IsSilent As Boolean)
Dim L() As FTNo: L = MthFTNoAy(A)
MdRmvFTNoAy A.Md, L
If Not IsSilent Then
    Debug.Print FmtQQ("MthRmv: Mth(?) of LinCnt(?) is deleted", MthDNm(A), FTNoAy_LinCnt(L))
End If
End Sub

Sub MthRpl(A As Mth, RplByLines$)
Dim F%: F = MthFmno(A)
MthRmv A
A.Md.InsertLines F, RplByLines
End Sub

Sub OyDo(Oy, DoFunNm$)
Dim O
For Each O In Oy
    Excel.Run DoFunNm, O ' DoFunNm cannot be like a Excel.Address (eg, A1, XX1)
Next
End Sub

Sub PjAddCls(A As VBProject, Nm$)
PjAddMbr A, Nm, vbext_ct_ClassModule
End Sub

Sub PjAddMbr(A As VBProject, Nm$, Ty As vbext_ComponentType, Optional IsGoMbr As Boolean)
If PjHasCmp(A, Nm) Then
    MsgBox FmtQQ("Cmp(?) exist in CurPj(?)", Nm, CurPjNm), , "M_A.ZAddMbr"
    Exit Sub
End If
Dim Cmp As VBComponent
Set Cmp = A.VBComponents.Add(Ty)
Cmp.Name = Nm
Cmp.CodeModule.InsertLines 1, "Option Explicit"
If IsGoMbr Then Shw_Mbr Nm
End Sub
Sub ZZ_PjAddRf()
PjAddRf Pj("QXls"), "QDta"
End Sub
Sub PjRmvRf(A As VBProject, RfNy0$)
AyDoPX DftNy(RfNy0), "PjRmvRf__X", A
PjSav A
End Sub
Sub PjAddRf(A As VBProject, RfNy0$)
AyDoPX DftNy(RfNy0), "PjAddRf__X", A
PjSav A
End Sub
Private Sub PjAddRf__X(A As VBProject, RfNm$)
If PjHasRfNm(A, RfNm) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfNm(?)", A.Name, RfNm)
    Exit Sub
End If
Dim RfFfn$: RfFfn = PjRfNm_RfFfn(A, RfNm)
If PjHasRfFfn(A, RfFfn) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfFfnNm(?)", A.Name, RfFfn)
    Exit Sub
End If
A.References.AddFromFile RfFfn
End Sub
Private Sub PjRmvRf__X(A As VBProject, RfNm$)
If PjHasRfNm(A, RfNm) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfNm(?)", A.Name, RfNm)
    Exit Sub
End If
Dim RfFfn$: RfFfn = PjRfNm_RfFfn(A, RfNm)
If PjHasRfFfn(A, RfFfn) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfFfnNm(?)", A.Name, RfFfn)
    Exit Sub
End If
A.References.AddFromFile RfFfn
End Sub

Sub PjCompile(A As VBProject)
PjGo A
SendKeys "%D{Enter}"
End Sub

Sub PjCrt_Fxa(A As VBProject, FxaNm$)
Dim F$
F = FxaNm_Fxa(FxaNm)
End Sub

Sub PjEnsCls(A As VBProject, ClsNm$)
PjEnsCmp A, ClsNm, vbext_ct_ClassModule
End Sub

Sub PjEnsCmp(A As VBProject, Nm$, Ty As vbext_ComponentType)
If PjHasCmp(A, Nm) Then Exit Sub
Dim Cmp As VBComponent
Set Cmp = A.VBComponents.Add(Ty)
Cmp.Name = Nm
Cmp.CodeModule.AddFromString "Option Explicit"
Debug.Print FmtQQ("PjEns_Cmp: Md(?) of Ty(?) is added in Pj(?) <===================================", Nm, CmpTy_Nm(Ty), A.Name)
End Sub

Sub PjEnsMd(A As VBProject, MdNm$)
PjEnsCmp A, MdNm, vbext_ct_StdModule
End Sub

Sub PjExport(A As VBProject)
Dim P$: P = PjSrcPth(A)
If P = "" Then
    Debug.Print FmtQQ("PjExport: Pj(?) does not have FileName", A.Name)
    Exit Sub
End If
PthClrFil P 'Clr SrcPth ---
FfnCpyToPth A.Filename, P, OvrWrt:=True
Dim I, Ay() As CodeModule
Ay = PjMbrAy(A)
If Sz(Ay) = 0 Then Exit Sub
For Each I In Ay
    MdExport CvMd(I)  'Exp each md --
Next
AyWrt PjRfLy(A), PjRfCfgFfn(A) 'Exp rf -----
End Sub
Function CmdBarOf_Std() As CommandBar
Set CmdBarOf_Std = CurVbe.CommandBars("Standard")
End Function
Function CmdBTonof_Std_Sav() As CommandBarButton
Dim I As CommandBarControl
For Each I In CmdBarOf_Std.Controls
    If IsPfx(I.Caption, "&Sav") Then Set CmdBTonof_Std_Sav = I: Exit Function
Next
Stop
End Function
Sub PjGo(A As VBProject)
Cls_Win
Dim Md As CodeModule
Set Md = PjFstMd(A)
If IsNothing(Md) Then
    Stop
    Exit Sub
End If
Md.CodePane.Show
SendKeys "%WV" ' Window SplitVertical
DoEvents
End Sub
Function PjTim(A As VBProject) As Date
PjTim = FfnTim(PjFfn(A))
End Function

Function FfnTim(A) As Date
FfnTim = FileDateTime(A)
End Function

Function PjFn$(A As VBProject)
PjFn = FfnFn(PjFfn(A))
End Function
Function DryToStr$(A)

End Function
Sub ZZ_PjSav()
PjSav CurPj
End Sub
Sub VbeSav(A As Vbe)
ItrDo A.VBProjects, "PjSav"
End Sub

Sub ZZ_VbeDmpIsSaved()
VbeDmpIsSaved CurVbe
End Sub
Sub VbeDmpIsSaved(A As Vbe)
Dim I As VBProject
For Each I In A.VBProjects
    Debug.Print I.Saved, I.BuildFileName
Next
End Sub
Function ItrPrpAy(A, PrpNm)
ItrPrpAy = ItrPrpAyInto(A, PrpNm, EmpAy)
End Function
Function ItrPrpAyInto(A, PrpNm, OInto)
Dim O: O = OInto: Erase O
Dim I
For Each I In A
    Push O, ObjPrp(I, PrpNm)
Next
ItrPrpAyInto = O
End Function
Sub ItrDo(A, DoFunNm$)
Dim I
For Each I In A
    Run DoFunNm, I
Next
End Sub
Sub PjSav(A As VBProject)
If A.Saved Then
    Debug.Print FmtQQ("PjSav: Pj(?) is already saved", A.Name)
    Exit Sub
End If
Dim Fn$: Fn = PjFn(A)
If Fn = "" Then
    Debug.Print FmtQQ("PjSav: Pj(?) needs saved first", A.Name)
    Exit Sub
End If
PjGo A
If ObjPtr(CurPj) <> ObjPtr(A) Then Stop
Dim B As CommandBarButton: Set B = CmdBTonof_Std_Sav
If Not StrIsEq(B.Caption, "&Save " & Fn) Then Stop
B.Execute
Debug.Print FmtQQ("PjSav: Pj(?) is not sure if saved <---------------", A.Name)
End Sub

Sub PjSrcPthBrw(A As VBProject)
PthBrw PjSrcPth(A)
End Sub

Sub PjSrt(A As VBProject)
Dim I
Dim Ny$(): Ny = AySrt(PjMd_and_Cls_Ny(A))
If Sz(Ny) = 0 Then Exit Sub
For Each I In Ny
    MdSrt PjMd(A, I)
Next
End Sub

Sub Pj_Gen_TstClass(A As VBProject)
If PjHasCmp(A, "Tst") Then
    CmpRmv PjCmp(A, "Tst")
End If
PjAddCls A, "Tst"
PjMd(A, "Tst").AddFromString Pj_TstClass_Bdy(A)
End Sub

Sub Pj_Gen_TstSub(A As VBProject)
Dim Ny$(): Ny = PjMd_and_Cls_Ny(A)
Dim N, M As CodeModule
For Each N In Ny
    Set M = A.VBComponents(N).CodeModule
    Md_Gen_TstSub M
Next
End Sub

'Function FfnRplExt$(Ffn, NewExt)
'FfnRplExt = FfnRmvExt(Ffn) & NewExt
'End Function
'Function FtDic(Ft) As Dictionary
'Set FtDic = Ly(FtLy(Ft)).Dic
'End Function
'Function FtLy(Ft) As String()
'Dim F%: F = FtOpnInp(Ft)
'Dim L$, O$()
'While Not EOF(F)
'    Line Input #F, L
'    Push O, L
'Wend
'Close #F
'FtLy = O
'End Function
'Function FtOpnApp%(Ft)
'Dim O%: O = FreeFile(1)
'Open Ft For Append As #O
'FtOpnApp = O
'End Function
'Function FtOpnInp%(Ft)
'Dim O%: O = FreeFile(1)
'Open Ft For Input As #O
'FtOpnInp = O
'End Function
'Function FtOpnOup%(Ft)
'Dim O%: O = FreeFile(1)
'Open Ft For Output As #O
'FtOpnOup = O
'End Function
Sub PthBrw(P)
Shell "Explorer """ & P & """", vbMaximizedFocus
End Sub

Sub PthClrFil(A)
Dim F
For Each F In PthFfnItr(A)
   FfnDlt F
Next
End Sub

Sub PthEns(P$)
If Fso.FolderExists(P) Then Exit Sub
MkDir P
End Sub

Sub Push(O, M)
Dim N&
    N = Sz(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub

Sub PushAy(OAy, Ay)
If Sz(Ay) = 0 Then Exit Sub
Dim I
For Each I In Ay
    Push OAy, I
Next
End Sub

Sub PushAyNoDup(OAy, Ay)
If Sz(Ay) = 0 Then Exit Sub
Dim I
For Each I In Ay
    PushNoDup OAy, I
Next
End Sub

Sub PushNoDup(O, M)
If Not AyHas(O, M) Then Push O, M
End Sub

Sub PushNonEmp(O, M)
If IsEmp(M) Then Exit Sub
Push O, M
End Sub

Sub PushObj(O, M)
If Not IsObject(M) Then Stop
Dim N&
    N = Sz(O)
ReDim Preserve O(N)
Set O(N) = M
End Sub

Sub PushObjAy(O, Oy)
If Sz(Oy) = 0 Then Exit Sub
Dim I
For Each I In Oy
    PushObj O, I
Next
End Sub

Sub RgVis(A As Range, Vis As Boolean)
If Vis Then A.Application.Visible = True
End Sub

Sub S1S2Ay_Brw(A() As S1S2)
AyBrw S1S2Ay_FmtLy(A)
End Sub

Sub SqSetRow(OSq, R&, Dr)
Dim J%
For J = 0 To UB(Dr)
    OSq(R, J + 1) = Dr(J)
Next
End Sub
Function ApItr(ParamArray Ap()) As Collection
Dim Av(): Av = Ap
Set ApItr = AyItr(Av)
End Function
Function StrLikItr(A, LikItr As Collection) As Boolean
Dim I
For Each I In LikItr
    If A Like I Then StrLikItr = True
Next
End Function
Sub StrBrw(A$)
Dim T$:
T = TmpFt
StrWrt A, T
Shell FmtQQ("code.cmd ""?""", T), vbMaximizedFocus
'Shell FmtQQ("notepad.exe ""?""", T), vbMaximizedFocus
End Sub

Sub StrWrt(A, Ft$, Optional IsNotOvrWrt As Boolean)
Fso.CreateTextFile(Ft, Overwrite:=Not IsNotOvrWrt).Write A
End Sub

Sub VbeClsWin(A As Vbe, Optional ExcptWinTyAy)
Dim W As VBIDE.Window
If IsEmpty(ExcptWinTyAy) Then
    ItrDoSub A.Windows, "Close"
    Exit Sub
End If
For Each W In A.Windows
    If Not AyHas(ExcptWinTyAy, W.Type) Then W.Close
Next
End Sub

Sub VbeExport(A As Vbe)
OyDo VbePjAy(A), "PjExport"
End Sub

Sub VbeSrcPthBrw(A As Vbe)
PthBrw VbeSrcPth(A)
End Sub

Sub VbeSrt(A As Vbe)
Dim I
For Each I In VbePjAy(A)
    PjSrt CvPj(I)
Next
End Sub

Sub VbeSrtRptBrw(A As Vbe)
AyBrw VbeSrtRptLy(A)
End Sub

Sub WbRfh(A As Workbook)
Dim Ws As Worksheet
For Each Ws In A.Worksheets
    WsRfh Ws
Next
End Sub

Sub WbRfhFcsv(A As Workbook, Fcsv$)
A.work
End Sub

Sub WbSetFcsv(A As Workbook, Fcsv$)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub

Sub WsRfh(A As Worksheet)
Dim L As ListObject, Qt As QueryTable
For Each L In A.ListObjects
    Set Qt = LoQt(L)
    If Not IsNothing(Qt) Then Qt.Refresh False
Next
Dim Q As QueryTable
For Each Q In A.QueryTables
    Q.Refresh False
Next
Dim P As PivotTable
For Each P In A.PivotTables
    P.RefreshTable
Next
End Sub

Sub XlsAddFxaNm(A As Excel.Application, FxaNm$)
Dim F$: F = FxaNm_Fxa(FxaNm)
If F = "" Then Exit Sub
A.AddIns.Add FxaNm_Fxa(FxaNm)
End Sub

Sub XlsVis(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub

Private Function DupMthFNyGp_CmpLy__1Hdr(OIx%, MthNm$, Cnt%) As String()
Dim O$(1)
O(0) = "================================================================"
Dim A$
    If OIx >= 0 Then A = FmtQQ("#DupFunNo(?) ", OIx): OIx = OIx + 1
O(1) = A + FmtQQ("DupFunNm(?) Cnt(?)", MthNm, Cnt)
DupMthFNyGp_CmpLy__1Hdr = O
End Function

Private Function DupMthFNyGp_CmpLy__2Sam(InclSam As Boolean, OSam%, DupMthFNyGp, LinesAy$()) As String()
If Not InclSam Then Exit Function
'{DupMthFNyGp} & {LinesAy} have same # of element
Dim O$()
Dim D$(): D = AyWhDup(LinesAy)
Dim J%, X$()
For J = 0 To UB(D)
    X = DupMthFNyGp_CmpLy__2Sam1(OSam, D(J), DupMthFNyGp, LinesAy)
    PushAy O, X
Next
DupMthFNyGp_CmpLy__2Sam = O
End Function

Private Function DupMthFNyGp_CmpLy__2Sam1(OSam%, Lines$, DupMthFNyGp, LinesAy$()) As String()
Dim A1$()
    If OSam > 0 Then
        Push A1, FmtQQ("#Sam(?) ", OSam)
        OSam = OSam + 1
    End If
Dim A2$()
    Dim J%
    For J = 0 To UB(LinesAy)
        If LinesAy(J) = Lines Then
            Push A2, "Shw """ & DupMthFNyGp(J) & """"
        End If
    Next
Dim A3$()
    A3 = LinesBoxLy(Lines)
DupMthFNyGp_CmpLy__2Sam1 = AyAddAp(A1, A2, A3)
End Function

Private Function DupMthFNyGp_CmpLy__3Syn(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim B$()
    Dim J%, I%
    Dim Lines
    For Each Lines In UniqLinesAy
        For I = 0 To UB(FunFNyGp)
            If Lines = LinesAy(I) Then
                Push B, FunFNyGp(I)
                Exit For
            End If
        Next
    Next
DupMthFNyGp_CmpLy__3Syn = AyMapPXSy(B, "FmtQQ", "Sync_Fun ""?""")
End Function

Private Function DupMthFNyGp_CmpLy__4Cmp(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim L2$() ' = From L1 with each element with MdDNm added in front
    ReDim L2(UB(UniqLinesAy))
    Dim Fnd As Boolean, DNm$, J%, Lines$, I%
    For J = 0 To UB(UniqLinesAy)
        Lines = UniqLinesAy(J)
        Fnd = False
        For I = 0 To UB(LinesAy)
            If LinesAy(I) = Lines Then
                DNm = FunFNyGp(I)
                L2(J) = DNm & vbCrLf & StrDup("-", Len(DNm)) & vbCrLf & Lines
                Fnd = False
                GoTo Nxt
            End If
        Next
        Stop
Nxt:
    Next
DupMthFNyGp_CmpLy__4Cmp = LinesAy_FmtLy(L2)
End Function

Private Function FunSync__1(A As Mth, Lines$) As Mth()
Dim Ny$(): Ny = FunNm_DupFunFNy(A.Nm)
Dim Ny1$(): Ny1 = AyRmvEle(Ny, MthFNm(A))
If Sz(Ny) <> Sz(Ny1) + 1 Then Stop
Dim O() As Mth, J%, M As Mth, L$
For J = 0 To UB(Ny1)
    Set M = MthFNm_Mth(Ny1(J))
    L = MthLines(M): If L = "" Then Stop
    If L <> Lines Then
        PushObj O, M
    End If
Next
If Sz(O) = 0 Then
    Debug.Print FmtQQ("FunSync: There are ?-Fun(?). All have same lines", Sz(Ny), MthDNm(A))
End If
FunSync__1 = O
End Function

Private Property Get ZZA()

End Property

Private Property Let ZZA(A)

End Property

Private Function ZZSrc() As String()
ZZSrc = MdSrc(CurMd)
End Function

Private Sub ZZZ_MdEndTrim()
Dim M As CodeModule: Set M = Md("ZZModule")
MdAppLines M, "  "
MdAppLines M, "  "
MdAppLines M, "  "
MdAppLines M, "  "
MdEndTrim M, ShwMsg:=True
Debug.Assert M.CountOfLines = 15
End Sub

Private Sub ZZZ_MthFTNoAy()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "ZZA")
Dim Act() As FTNo: Act = MthFTNoAy(M)
Debug.Assert Sz(Act) = 2
Debug.Assert Act(0).Fmno = 5
Debug.Assert Act(0).Tono = 7
Debug.Assert Act(1).Fmno = 13
Debug.Assert Act(1).Tono = 15
End Sub

Private Sub ZZZ_MthRmv()
Dim M As CodeModule: Set M = Md("ZZModule")
Dim M1 As Mth, M2 As Mth
Set M1 = Mth(M, "ZZRmv1")
Set M2 = Mth(M, "ZZRmv2")
MdAppLines M, RplVBar("Function ZZRmv1()||End Property||Function ZZRmv2()|End Function||Property Let ZZRmv1(V)|End Property")
MthRmv M1
MthRmv M2
MdEndTrim M
Debug.Assert M.CountOfLines = 15
End Sub

Private Sub ZZZ_WbSetFcsv()
Dim Wb As Workbook
Set Wb = WbOf_Mth
Debug.Print WbTxtCnStr(Wb)
WbSetFcsv Wb, "C:\ABC.CSV"
Debug.Assert WbTxtCnStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub

Private Sub ZZZ_WbTxtCnCnt()
Dim O As Workbook: Set O = WbOf_Mth
Debug.Assert WbTxtCnCnt(O) = 1
O.Close
End Sub

Private Sub ZZ_LinesAy_FmtLy()
Dim A$()
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf|sdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdf|skldfjdf|lskdf|slkdjf|sdlf||")
Push A, RplVBar("ksdjlfdf|sdklfjsdfdsfdsf|skldsdffjdf")
AyDmp LinesAy_FmtLy(A)
End Sub

Private Sub ZZ_MdCmpByNm()
MdCmpByNm "QTool.G_Tool", "QVb.M_Ay"
End Sub

Private Sub ZZ_MdDicOfMthNmzzzMthLines()
DicBrw MdDicOfMthNmzzzMthLines(CurMd)
End Sub

Private Sub ZZ_MdMthNyOfInproper()
AyDmp MdMthNyOfInproper(Md("QDta.M_Ay"))
End Sub

Private Sub ZZ_Md_FunNm_z_ProperMdNm_Brw()
Md_FunNm_z_ProperMdNm_Brw CurMd
End Sub

Private Sub ZZ_MthLin_MthKey()
Dim Ay1$(): Ay1 = SrcMthLinAy(CurSrc)
Dim Ay2$(): Ay2 = AyMapSy(Ay1, "MthLin_MthKey")
S1S2Ay_Brw AyAB_S1S2Ay(Ay2, Ay1)
End Sub

Private Sub ZZ_MthLin_MthKey_1()
Const A$ = "Function ZZA()"
Debug.Print MthLin_MthKey(A, IsWrap:=True)
End Sub

Private Sub ZZ_FunNm_Cmp()
FunNm_Cmp "FfnDlt"
End Sub

Private Sub ZZ_SrcMthBrkAy()
Dim A() As MthBrk: A = SrcMthBrkAy(CurSrc)
AyBrw OyToStrSy(A)
End Sub

Private Sub ZZ_SrcDclLinCnt()
Dim B1$(): B1 = CurSrc
Dim B2$(): B2 = SrcSrtedLy(B1)
Dim A1%: A1 = SrcDclLinCnt(B1)
Dim A2%: A2 = SrcDclLinCnt(SrcSrtedLy(B1))
End Sub

Private Sub ZZ_SrcDicOfMthNmzzzMthLines()
'Dim A As Dictionary: Set A = SrcDicOfMthNmzzzMthLines(CurSrc)
DicBrw SrcDicOfMthNmzzzMthLines(CurSrc)
End Sub

Private Sub ZZ_SrcSrtRptLy()
AyBrw SrcSrtRptLy(CurSrc, "Pj", "Md")
End Sub

Private Sub ZZ_SrcSrtedBdyLines()
StrBrw SrcSrtedBdyLines(CurSrc)
End Sub

Private Sub ZZ_VbeDupFunCmpLy()
AyBrw VbeDupFunCmpLy(CurVbe)
End Sub
Sub A1()
'...
End Sub
Private Sub ZZ_VbeFunFNy()
AyBrw VbeFunFNy(CurVbe, ExclFunNy0:="ZZZ__Tst")
End Sub

Private Sub ZZ_VbeFunPfxAy()
AyDmp VbeFunPfxAy(CurVbe)
End Sub

Private Sub ZZ_XlsAddFxaNm()
XlsAddFxaNm Application, "QIde0"
End Sub

Function DftFun(FunDNm0$) As Mth
If FunDNm0 = "" Then
    Dim M As Mth
    Set M = CurMth
    If IsFun(M) Then
        Set DftFun = M
    End If
Else
End If
Stop '
End Function

Function IsMthDNm(Nm) As Boolean
IsMthDNm = Sz(Split(Nm, ".")) = 3
End Function

Function IsMthFNm(Nm) As Boolean
Dim P%: P = InStr(Nm, ":"): If P = 0 Then Exit Function
IsMthFNm = InStr(Nm, ".") > P
End Function

