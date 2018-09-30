Attribute VB_Name = "VbAy"
Option Explicit

Function AyFlat(AyOfAy)
If Sz(AyOfAy) = 0 Then Exit Function
AyFlat = AyOfAy(0)
Erase AyFlat
Dim X
For Each X In AyOfAy
    PushAy AyFlat, X
Next
End Function


Function AyAsg(A, ParamArray OAp())
Dim Av: Av = OAp
Dim J%
For J = 0 To UB(Av)
    OAp(J) = A(J)
Next
End Function
Function AyTrim(A) As String()
Dim X
For Each X In AyNz(A)
    Push AyTrim, Trim(X)
Next
End Function
Function AyWrpPad(A, W%) As String() ' Each Itm of Ay[A] is padded to line with AyWdt(A).  return all padded lines as String()
Dim O$(), X, I%
ReDim O(0)
For Each X In AyNz(A)
    If Len(O(I)) + Len(X) < W Then
        O(I) = O(I) & X
    Else
        PushI O, X
        I = I + 1
    End If
Next
AyWrpPad = O
End Function
Function AyAddCommaSpcSfxExptLas(A) As String()
Dim X, J, U%
U = UB(A)
For Each X In AyNz(A)
    If J = U Then
        Push AyAddCommaSpcSfxExptLas, X
    Else
        Push AyAddCommaSpcSfxExptLas, X & ", "
    End If
    J = J + 1
Next
End Function
Function AyDistIdCntDic(A) As Dictionary
'Type DistIdCntDic = Map Val [Id,Cnt]
Dim X, O As New Dictionary, J&, IdCnt()
For Each X In AyNz(A)
    If Not O.Exists(X) Then
        O.Add X, Array(J, 1)
        J = J + 1
    Else
        IdCnt = O(X)
        O(X) = Array(IdCnt(0), IdCnt(1) + 1)
    End If
Next
Set AyDistIdCntDic = O
End Function
Function AySeqCntDic(A) As Dictionary 'The return dic of key=AyEle pointing to Long(1) with Itm0 as Seq# and Itm1 as Cnt
Dim S&, O As New Dictionary, L&(), X
For Each X In AyNz(A)
    If O.Exists(X) Then
        L = O(X)
        L(1) = L(1) + 1
        O(X) = L
    Else
        ReDim L(1)
        L(0) = S
        L(1) = 1
        O.Add X, L
    End If
Next
Set AySeqCntDic = O
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
Function AyabFmt(A, B) As String()
AyabFmt = S1S2AyFmt(AyabS1S2Ay(A, B))
End Function
Function AyabS1S2Ay(A, B) As S1S2()
Dim U&: U = UB(A)
If U <> UB(B) Then Stop
Dim O() As S1S2
ReDim O(U)
Dim J&
For J = 0 To U
    Set O(J) = S1S2(A(J), B(J))
Next
AyabS1S2Ay = O
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
Function AySel(A, M) As Boolean
If Sz(A) = 0 Then AySel = True: Exit Function
AySel = AyHas(A, M)
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
Function AyIns2(A, X1, X2, Optional At&)
Dim O
O = AyReSzAt(A, At, 2)
Asg X1, O(At)
Asg X2, O(At + 1)
AyIns2 = O
End Function
Function AyIsAllEleEq(A) As Boolean
If Sz(A) = 0 Then AyIsAllEleEq = True: Exit Function
Dim J&
For J = 1 To UB(A)
    If A(0) <> A(J) Then Exit Function
Next
AyIsAllEleEq = True
End Function
Function AyIx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then AyIx = J: Exit Function
Next
AyIx = -1
End Function
Function AyRmvLasEle(A)
Dim O
O = A
ReDim Preserve O(UB(A) - 1)
AyRmvLasEle = O
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
Function AyMapPXInto(A, PX$, P, OIntoAy)
'MapPXFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
' or "Mth"
Dim O: O = OIntoAy: Erase O
Dim X
If Sz(A) > 0 Then
    For Each X In A
        Push O, Run(PX, P, X)
    Next
End If
AyMapPXInto = O
End Function
Function AyMapXPInto(A, XP$, P, OInto)
'MapXPFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
Dim O, X
O = OInto
Erase O
For Each X In AyNz(A)
    Push O, Run(XP, X, P)
Next
AyMapXPInto = O
End Function
Function AyMapXAB(Ay, XAB$, A, B)
AyMapXAB = AyMapXABInto(Ay, XAB, A, B, EmpSy)
End Function
Function AyMapXABInto(Ay, XAB$, A, B, OInto)
'MapXPFunNm cannot be Excel-like-Function-Name, eg A2, AA2, (cell-address)
Dim O, X
O = OInto
Erase O
For Each X In AyNz(A)
    Push O, Run(XAB, X, A, B)
Next
AyMapXABInto = O
End Function
Function AyMapPXSy(A, PX$, P) As String()
AyMapPXSy = AyMapPXInto(A, PX, P, EmpSy)
End Function
Function AyMapXPSy(A, XP$, P) As String()
AyMapXPSy = AyMapXPInto(A, XP, P, EmpSy)
End Function
Function AyMapXP(A, XP$, P) As Variant()
AyMapXP = AyMapXPInto(A, XP, P, EmpAy)
End Function
Function AyMapPX(A, XP$, P) As Variant()
AyMapPX = AyMapPXInto(A, XP, P, EmpAy)
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
Function AyabDic(A1, A2) As Dictionary
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
Set AyabDic = O
End Function
Function AyRgH(Ay, At As Range) As Range
Set AyRgH = SqRg(AySqH(Ay), At)
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
'If W > 50 Then
'    For Each I In A
'    If Len(I) > 50 Then Debug.Print I: Stop
'    Next
'End If
End Function
Function AyWhDist(A)
If Sz(A) = 0 Then AyWhDist = A: Exit Function
Dim O: O = A: Erase O
Dim I
For Each I In A
    PushNoDup O, I
Next
AyWhDist = O
End Function
Function AyWhDup(A)
If Sz(A) = 0 Then AyWhDup = A: Exit Function
Dim O: O = A: Erase O
Dim Dr
For Each Dr In AyNz(AyCntDry(A))
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
Function AyWhExlAy(A, ExlAy$()) As String()
If Sz(ExlAy) = 0 Then AyWhExlAy = AySy(A): Exit Function
Dim X
For Each X In A
    If Not IsInLikAy(X, ExlAy) Then PushI AyWhExlAy, X
Next
End Function
Function AyWhExl(A, Exl$) As String()
AyWhExl = AyWhExlAy(A, SslSy(Exl))
If Sz(A) = 0 Then Exit Function
End Function
Function AyMid(A, Fm, Optional L)
Dim O: O = A: Erase O
Dim J&, E&
If L = 0 Then E = UB(A) Else E = L + Fm - 1
For J = Fm To E
    Push O, A(J)
Next
AyMid = O
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
Function AyWhNm(A, B As WhNm) As String()
AyWhNm = AyWhExlAy(AyWhRe(A, B.Re), B.ExlAy)
End Function
Function AyWhRe(A, Re As RegExp) As String()
If IsNothing(Re) Then AyWhRe = AySy(A): Exit Function
Dim X
For Each X In AyNz(A)
    If Re.Test(X) Then PushI AyWhRe, X
Next
End Function
Function AyWhPatnExl(A, Patn$, Exl$) As String()
AyWhPatnExl = AyWhExl(AyWhPatn(A, Patn), Exl)
End Function
Function AyWhPatn(A, Patn$) As String()
If Sz(A) = 0 Then Exit Function
If Patn = "" Or Patn = "." Then AyWhPatn = AySy(A): Exit Function
Dim X, R As RegExp
Set R = Re(Patn)
For Each X In A
    If R.Test(X) Then Push AyWhPatn, X
Next
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
Function AyPredAllTrue(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPredAllTrue = ItrPredAllTrue(A, Pred)
End Function
Function AyPred_IsAllFalse(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_IsAllFalse = ItrPredAllFalse(A, Pred)
End Function
Function AyPred_HasSomTrue(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPred_HasSomTrue = ItrPredSomTrue(A, Pred)
End Function
Function AyPredSomFalse(A, Pred$) As Boolean
If Sz(A) = 0 Then Exit Function
AyPredSomFalse = ItrPredSomFalse(A, Pred)
End Function
Function AyabWs(A, B, Optional N1$ = "Ay1", Optional N2$ = "Ay2", Optional LoNm$ = "AyAB") As Worksheet
Dim N&, AtA1 As Range, R As Range
N = Sz(A)
If N <> Sz(B) Then Stop
Set AtA1 = NewA1

AyRgH Array(N1, N2), AtA1
AyRgV A, AtA1.Range("A2")
AyRgV B, AtA1.Range("B2")
Set R = RgRCRC(AtA1, 1, 1, N + 1, 2)
RgLo R, LoNm
Set AyabWs = AtA1.Parent
End Function
Function AyRgV(A, At As Range) As Range
Set AyRgV = SqRg(AySqV(A), At)
End Function
Function AyItmCnt%(A, M)
If Sz(A) = 0 Then Exit Function
Dim O%, X
For Each X In A
    If X = M Then O = O + 1
Next
AyItmCnt = O
End Function
Function AyGpCntDry(A) As Variant()
If Sz(A) = 0 Then Exit Function
Dim Dup, O(), X, T&, Cnt&
Dup = AyWhDist(A)
For Each X In AyNz(Dup)
    Cnt = AyItmCnt(A, X)
    Push O, Array(X, AyItmCnt(A, X))
    T = T + Cnt
Next
Push O, Array("~Tot", T)
AyGpCntDry = O
End Function
Function AyGpCntDryWhDup(A) As Variant()
AyGpCntDryWhDup = DryWhColGt(AyGpCntDry(A), 1, 1)
End Function
Function AyGpCntFmt(A) As String()
AyGpCntFmt = DryFmtss(AyGpCntDry(A))
End Function
Function AyAdd1(A)
AyAdd1 = AyAddN(A, 1)
End Function
Function AyAddN(A, N%)
If Sz(A) = 0 Then Exit Function
Dim O, U&
O = A
Dim J&
For J = 0 To U
    O(J) = A(J) + N
Next
AyAddN = O
End Function
Function AyIntAy(A) As Integer()
If Sz(A) = 0 Then Exit Function
Dim I, O%(), J&
ReDim O%(UB(A))
For Each I In A
    O(J) = I
    J = J + 1
Next
AyIntAy = O
End Function
Function AyMin(A)
Dim O, J&
If Sz(A) = 0 Then Exit Function
O = A(0)
For J = 1 To UB(A)
    If A(J) < O Then O = A(J)
Next
AyMin = O
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
Function AyAdd(A, B)
Dim O, X
O = A
For Each X In AyNz(B)
    Push O, X
Next
AyAdd = O
End Function
Function AyCln(A)
Dim O
O = A
Erase O
AyCln = O
End Function
Function AyReSz(A, SzAy) ' return Ay as [A] with sz as [SzAy]
If Sz(SzAy) = 0 Then AyReSz = A: Exit Function
Dim O
O = A
ReDim O(UB(SzAy))
AyReSz = O
End Function
Sub AyBrw(A, Optional Fnn$)
Dim T$
T = TmpFt("Brw", Fnn)
AyWrt A, T
FtBrw T
End Sub
Sub AyDmp(A)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    D I
Next
End Sub
Sub AyDoPX(A, DoMthNm$, P)
If Sz(A) = 0 Then Exit Sub
Dim X
For Each X In A
    Run DoMthNm, P, X
Next
End Sub
Sub AyDoXP(A, XP$, P)
If Sz(A) = 0 Then Exit Sub
Dim X
For Each X In A
    Run XP, X, P
Next
End Sub
Sub AyDo(A, DoMthNm$)
If Sz(A) = 0 Then Exit Sub
Dim I
For Each I In A
    Run DoMthNm, I
Next
End Sub
Sub AyWrt(A, FT$)
StrWrt JnCrLf(A), FT
End Sub
Function AyIxAy(A, B) As Long()
AyIxAy = AyIxAyInto(A, B, EmpLngAy)
End Function
Function AyIxAyI(A, B) As Integer()
AyIxAyI = AyIxAyInto(A, B, EmpIntAy)
End Function
Function AyIxAyInto(A, B, OIntoAy)
Dim J&, U&, O
O = OIntoAy
Erase O
U = UB(B)
ReDim O(U)
For J = 0 To U
    O(J) = AyIx(A, B(J))
Next
AyIxAyInto = O
End Function
Function IsAyOfStr(A) As Boolean
If Not IsVy(A) Then Exit Function
Dim X
For Each X In AyNz(A)
    If Not IsStr(X) Then Exit Function
Next
IsAyOfStr = True
End Function
Function IsVy(A) As Boolean
IsVy = VarType(A) = vbArray + vbVariant
End Function
Function IsAyOfAy(A) As Boolean
If Not IsVy(A) Then Exit Function
Dim X
For Each X In AyNz(A)
    If Not IsArray(X) Then Exit Function
Next
IsAyOfAy = True
End Function
Function AyMapFlat(Ay, ItmAyFun$)
AyMapFlat = AyMapFlatInto(Ay, ItmAyFun, EmpAy)
End Function
Function AyMapFlatInto(Ay, ItmAyFun$, OIntoAy)
Dim O, J&, M
O = OIntoAy: Erase O
For J = 0 To UB(Ay)
    M = Run(ItmAyFun, Ay(J))
    PushAy O, M
Next
AyMapFlatInto = O
End Function
Function AyReSzAt(A, At&, Optional Cnt& = 1)
If Cnt < 1 Then Stop
Dim O, U&, J&, F&, T&
O = A
U = UB(A)
'----------
Dim NU&, FOS&, TOS&, UM&
NU = U + Cnt
FOS = U
TOS = NU
UM = U - At
ReDim Preserve O(NU)
For J = 0 To UM
    F = FOS - J
    T = TOS - J
    Asg O(F), O(T)
    O(F) = Empty
Next
AyReSzAt = O
End Function
Function AyInsAy(A, B, Optional At&)
Dim O, NA&, NB&, J&
NA = Sz(A)
NB = Sz(B)
O = AyReSzAt(A, At, NB)
For J = 0 To NB - 1
    Asg B(J), O(At + J)
Next
AyInsAy = O
End Function
Function AyPop(A) As Variant()
AyPop = Array(AyLasEle(A), AyRmvLasEle(A))
End Function
Function AyNz(A)
If Sz(A) = 0 Then Set AyNz = New Collection: Exit Function
AyNz = A
End Function
Function AyFmt(A, ColBrkss$) As String()
AyFmt = DryFmtss(AyColBrkssDry(A, ColBrkss))
End Function
Function AyColBrkssDry(A, ColBrkss$) As Variant()
Dim Lin, Ay$()
Ay = SslSy(ColBrkss)
For Each Lin In AyNz(A)
    PushI AyColBrkssDry, LinBrkssDr(Lin, Ay)
Next
End Function
Sub ZZ_AyIns()

End Sub
