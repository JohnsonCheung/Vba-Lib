Attribute VB_Name = "M_Ay"
Option Explicit

Function AyAdd(Ay1, Ay2)
Dim O: O = Ay1
PushAy O, Ay2
AyAdd = O
End Function

Function AyAddAp(Ay, ParamArray Itm_or_Ay_Ap())
Dim Av(): Av = Itm_or_Ay_Ap
Dim O, I
O = Ay
For Each I In Av
    If IsArray(I) Then
        PushAy O, I
    Else
        Push O, I
    End If
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
If AyIsEmp(Ay) Then Exit Function
Dim W%: W = AyWdt(Ay)
Dim O$(), I
For Each I In Ay
    Push O, AlignL(I, W)
Next
AyAlignL = O
End Function

Function AyCellSy(Ay, Optional ShwZer As Boolean) As String()
Dim O$(), I, J&, U&
U = UB(Ay)
ReSz O, U
For Each I In Ay
    O(J) = VarCellStr(I)
    J = J + 1
Next
AyCellSy = O
End Function

Function AyC1Dry(Ay, C) As Variant()
'C1Dry is Dry with 2 column and Col1 is const
Dim U&, J&
U = UB(Ay)
Dim O()
ReSz O, U
For J = 0 To U
    O(J) = Array(C, Ay(J))
Next
AyC1Dry = O
End Function

Function AyC2Dry(Ay, C) As Variant()
'C2Dry is Dry with 2 column and Col1 is const
Dim U&, J&
U = UB(Ay)
Dim O()
ReSz O, U
For J = 0 To U
    O(J) = Array(Ay(J), C)
Next
AyC2Dry = O
End Function

Function AyDic(Ay, Optional V = True) As Dictionary
Dim O As New Dictionary, I
If Not AyIsEmp(Ay) Then
    For Each I In Ay
        O.Add I, V
    Next
End If
Set AyDic = O
End Function

Function AyDry(Ay) As Variant()
Dim O(), J&
Dim U&: U = UB(Ay)
ReSz O, U
For J = 0 To U
    O(J) = Array(Ay(J))
Next
AyDry = O
End Function

Function AyDupAy(Ay)
'Return Array of element of {Ay} for which has 2 or more value in {Ay}
Dim OAy: OAy = Ay: Erase OAy
If Not AyIsEmp(Ay) Then
    Dim Uniq: Uniq = OAy
    Dim V

    For Each V In Ay
        If AyHas(Uniq, V) Then
            Push OAy, V
        Else
            Push Uniq, V
        End If
    Next
End If
AyDupAy = OAy
End Function

Function AyEqChk(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz(Ay1), Sz(Ay2)): GoTo X
If AyIsEmp(Ay1) Then Exit Function
Dim O1$()
    Dim A2: A2 = Ay2
    Dim J&, ReachLimit As Boolean
    Dim Cnt%
    For J = 0 To U
        If Ay1(J) <> Ay2(J) Then
            Push O1, FmtQQ("[?]-th Ele is diff: ?[?]<>?[?]", Ay1Nm, Ay2Nm, Ay1(J), Ay2(J))
            Cnt = Cnt + 1
        End If
        If Cnt > 10 Then
            ReachLimit = True
            Exit For
        End If
    Next
If M_Is.IsEmp(O1) Then Exit Function
Dim O2$()
    Push O2, FmtQQ("Array [?] and [?] both having size[?] have differnt element(s):", Ay1Nm, Ay2Nm, Sz(Ay1))
    If ReachLimit Then
        Push O2, FmtQQ("At least [?] differences:", Sz(O1))
    End If
PushAy O, O2
PushAy O, O1
X:
Push O, FmtQQ("Ay-[?]:", Ay1Nm)
PushAy O, AyQuote(Ay1, "[]")
Push O, FmtQQ("Ay-[?]:", Ay2Nm)
PushAy O, AyQuote(Ay2, "[]")
AyEqChk = O
End Function

Function AyBrkInto3Ay(A, Fmix&, Toix&) As Variant()
Dim O(2)
O(0) = AyWhFmTo(A, 0, Fmix - 1)
O(1) = AyWhFmTo(A, Fmix, Toix)
O(2) = AyWhFm(A, Toix + 1)
AyBrkInto3Ay = O
End Function
Function AyWhFmTo(A, Fmix, Toix)
Dim O: O = A: Erase O
Dim J&
For J = Fmix To Toix
    Push O, A(J)
Next
AyWhFmTo = O
End Function

Function AyFstNEle(A, N&)
Dim O: O = A
ReDim Preserve O(N - 1)
AyFstNEle = O
End Function

Function AyGpDry(Ay) As Variant()
If AyIsEmp(Ay) Then Exit Function
Dim O(), I
For Each I In Ay
    AyGpDry__Upd O, I
Next
AyGpDry = O
End Function

Function AyHas(Ay, Itm) As Boolean
If AyIsEmp(Itm) Then Exit Function
Dim I
For Each I In Ay
    If I = Itm Then AyHas = True: Exit Function
Next
End Function

Function AyHasDupEle(Ay) As Boolean
If AyIsEmp(Ay) Then Exit Function
Dim Pool: Pool = Ay: Erase Pool
Dim I
For Each I In Ay
    If AyHas(Pool, I) Then AyHasDupEle = True: Exit Function
    Push Pool, I
Next
End Function

Function AyHasNegOne(Ay) As Boolean
Dim V
If AyIsEmp(Ay) Then Exit Function
For Each V In Ay
    If V = -1 Then AyHasNegOne = True: Exit Function
Next
End Function

Function AyHasSubAy(Ay, SubAy) As Boolean
If AyIsEmp(Ay) Then Exit Function
If AyIsEmp(SubAy) Then ErPm
Dim I
For Each I In SubAy
    If Not AyHas(Ay, I) Then Exit Function
Next
End Function

Function AyIncNForEachEle(Ay, Optional N& = 1)
Dim O: O = Ay
Dim J&
For J = 0 To UB(Ay)
    O(J) = O(J) + N
Next
AyIncNForEachEle = O
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

Function AyIntersect(Ay1, Ay2)
Dim O: O = Ay1: Erase O
If AyIsEmp(Ay1) Then GoTo X
If AyIsEmp(Ay2) Then GoTo X
Dim V
For Each V In Ay1
    If AyHas(Ay2, V) Then O.Push V
Next
X:
AyIntersect = O
End Function

Function AyIsAllEleHasPfx(A, Pfx$) As Boolean
If AyIsEmp(A) Then Exit Function
Dim I
For Each I In A
   If Not HasPfx(I, Pfx) Then Exit Function
Next
AyIsAllEleHasPfx = True
End Function

Function AyIsAllEleHasVal(Ay) As Boolean
If AyIsEmp(Ay) Then Exit Function
Dim I
For Each I In Ay
    If M_Is.IsEmp(I) Then Exit Function
Next
AyIsAllEleHasVal = True
End Function

Function AyIsAllEq(Ay) As Boolean
If AyIsEmp(Ay) Then AyIsAllEq = True: Exit Function
Dim T: T = Ay(0)
Dim J&
For J = 1 To UB(Ay)
    If Ay(J) = T Then Exit Function
Next
AyIsAllEq = True
End Function

Function AyIsAllStr(Ay) As Boolean
If Sz(Ay) = 0 Then Exit Function
Dim K
For Each K In Ay
    If Not IsStr(K) Then Exit Function
Next
AyIsAllStr = True
End Function

Function AyIsEmp(V) As Boolean
AyIsEmp = Sz(V) = 0
End Function

Function AyIsEq(A1, A2) As Boolean
Dim U&: U = UB(A1): If U <> UB(A2) Then Exit Function
Dim J&
For J = 0 To U
   If A1(J) <> A2(J) Then Exit Function
Next
AyIsEq = True
End Function

Function AyIsEqSz(Ay, B) As Boolean
AyIsEqSz = Sz(Ay) = Sz(B)
End Function

Function AyIsSamSz(Ay1, Ay2) As Boolean
AyIsSamSz = Sz(Ay1) = Sz(Ay2)
End Function

Function AyIsSrt(Ay) As Boolean
Dim J&
For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Function
Next
AyIsSrt = True
End Function

Function AyIx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then AyIx = J: Exit Function
Next
AyIx = -1
End Function

Function AyIxAy(Ay, SubAy, Optional ChkNotFound As Boolean, Optional SkipNotFound As Boolean) As Long()
If AyIsEmp(SubAy) Then Exit Function
Dim O&()
Dim U&: U = UB(SubAy)
Dim J&, Ix&
If SkipNotFound Then
    For J = 0 To U
        Ix = AyIx(Ay, SubAy(J))
        If Ix >= 0 Then
            Push O, Ix
        End If
    Next
Else
    ReDim O(U)
    For J = 0 To U
        O(J) = AyIx(Ay, SubAy(J))
    Next
End If
If Not SkipNotFound And ChkNotFound Then
    AyIxAy__ChkNotFound O, Ay, SubAy
End If
AyIxAy = O
End Function

Function AyLasEle(Ay)
AyLasEle = Ay(UB(Ay))
End Function

Function AyMap(Ay, MthNm$, ParamArray Ap()) As Variant()
If AyIsEmp(Ay) Then Exit Function
Dim Av(): Av = Ap
Av = AyIns(Av)
Dim I, J&
Dim O()
Dim U&: U = UB(Ay)
    ReDim O(U)
For Each I In Ay
    Asg I, Av(0)
    Asg RunAv(MthNm, Av), O(J)
    J = J + 1
Next
AyMap = O
End Function

Function AyMapAsgAy(Ay, OAy, MthNm$, ParamArray Ap())
If AyIsEmp(Ay) Then Exit Function
Dim Av(): Av = Ap
Av = AyIns(Av)
Dim I, J&
Dim O
O = OAy
Erase O
Dim U&: U = UB(Ay)
    ReDim O(U)
For Each I In Ay
    Asg I, Av(0)
    Asg RunAv(MthNm, Av), O(J)
    J = J + 1
Next
AyMapAsgAy = O
End Function

Function AyMapInto(A, MapFunNm$, OIntoAy)
Dim O: O = OIntoAy: Erase OIntoAy
Dim I
If Sz(A) > 0 Then
    For Each I In A
        Push O, Run(MapFunNm, I)
    Next
End If
AyMapInto = O
End Function

Function AyMapSy(A, MapFunNm$) As String()
AyMapSy = AyMapInto(A, MapFunNm, EmpSy)
End Function

Function AyMapLngAy(Ay, MapMthNm$) As Long()
AyMapLngAy = AyMapInto(Ay, MapMthNm, EmpLngAy)
End Function

Function AyMax(A)
Dim O: O = A(0)
Dim J&
For J = 1 To UB(A)
    O = Max(O, A(J))
Next
AyMax = O
End Function

Function AyMinus(Ay1, Ay2)
If AyIsEmp(Ay1) Then AyMinus = Ay1: Exit Function
Dim O: O = Ay1: Erase O
Dim mAy2: mAy2 = Ay2
Dim V
For Each V In Ay1
    If AyHas(mAy2, V) Then
        mAy2 = AyRmvEle(mAy2, V)
    Else
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

Function AyNoDupAy(Ay)
Dim O: O = Ay
Erase O
Dim I
If AyIsEmp(Ay) Then AyNoDupAy = O: Exit Function
For Each I In Ay
    PushNoDup O, I
Next
AyNoDupAy = O
End Function

Function AyQuote(Ay, QuoteStr$) As String()
If AyIsEmp(Ay) Then Exit Function
Dim O$(), U&
    U = UB(Ay)
    ReDim O(U)
    Dim J&
    Dim Q1$, Q2$
    S1S2_Asg BrkQuote(QuoteStr), Q1, Q2
    For J = 0 To U
        O(J) = Q1 & Ay(J) & Q2
    Next
AyQuote = O
End Function

Function AyQuoteDbl(Ay) As String()
AyQuoteDbl = AyQuote(Ay, """")
End Function

Function AyQuoteSng(Ay) As String()
AyQuoteSng = AyQuote(Ay, "'")
End Function

Function AyQuoteSqBkt(Ay) As String()
AyQuoteSqBkt = AyQuote(Ay, "[]")
End Function

Function AyRTrim(Ay) As String()
If AyIsEmp(Ay) Then Exit Function
Dim O$(), I
For Each I In Ay
    Push O, RTrim(I)
Next
AyRTrim = O
End Function

Function AyReOrd(Ay, PartialIxAy&())
Dim I&()
    I = PartialIxAy_CompleteIxAy(PartialIxAy, UB(Ay))
Dim O
    O = Ay: Erase O
    Dim J&
    For J = 0 To UB(I)
        Push O, Ay(I(J))
    Next
AyReOrd = O
End Function

Function AyRmvEle(Ay, Ele)
Dim Ix&: Ix = AyIx(Ay, Ele): If Ix = -1 Then AyRmvEle = Ay: Exit Function
AyRmvEle = AyRmvEleAt(Ay, AyIx(Ay, Ele))
End Function

Function AyRmvEleAt(Ay, Optional At&)
AyRmvEleAt = AyWhExclAtCnt(Ay, At)
End Function

Function AyRmvEmpEle(Ay)
If AyIsEmp(Ay) Then AyRmvEmpEle = Ay: Exit Function
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If Not IsEmp(I) Then Push O, I
Next
AyRmvEmpEle = O
End Function

Function AyRmvEmpEleAtEnd(Ay)
Dim LasU&, U&
Dim O: O = Ay
For LasU = UB(Ay) To 0 Step -1
    If Not IsEmp(O(LasU)) Then
        Exit For
    End If
Next
If LasU = -1 Then
    Erase O
Else
    ReDim Preserve O(LasU)
End If
AyRmvEmpEleAtEnd = O
End Function

Function AyRmvFmTo(Ay, FmTo As FmTo)
Dim O
    O = Ay
    If Not FmTo_IsVdt(FmTo) Or AyIsEmp(Ay) Then
        Dim FmI&, ToI&
        FmI = FmTo.Fmix
        ToI = FmTo.Toix
        Dim I&, J&, U&
        U = UB(Ay)
        I = 0
        For J = ToI + 1 To U
            O(FmI + I) = O(J)
            I = I + 1
        Next
        ReDim Preserve O(U - FmTo_Cnt(FmTo))
    End If
AyRmvFmTo = O
End Function

Function AyRmvFstChr(A) As String()
AyRmvFstChr = AyMapSy(A, "RmvFstChr")
End Function

Function AyRmvFstEle(Ay)
AyRmvFstEle = AyRmvEleAt(Ay)
End Function

Function AyRmvLasChr(A) As String()
AyRmvLasChr = AyMapSy(A, "RmvLasChr")
End Function

Function AyRmvLasEle(Ay)
AyRmvLasEle = AyRmvEleAt(Ay, UB(Ay))
End Function

Function AyRmvPfx(Ay, Pfx) As String()
If AyIsEmp(Ay) Then Exit Function
Dim U&: U = UB(Ay)
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = RmvPfx(Ay(J), Pfx)
Next
AyRmvPfx = O
End Function

Function AyRpl(Ay, Fmix&, Toix&, AySeg)
Dim A()
    A = AyBrkInto3Ay(Ay, Fmix, Toix)
Dim O
    O = Ay(0): Erase O
    PushAy O, AySeg
    PushAy O, Ay(2)
AyRpl = O
End Function

Function AyShift(OAy)
AyShift = OAy(0)
OAy = AyRmvFstEle(OAy)
End Function

Function AySqH(Ay) As Variant()
If AyIsEmp(Ay) Then Exit Function
Dim O(), C%
ReDim O(1 To 1, 1 To Sz(Ay))
C = 0
Dim V
For Each V In Ay
    C = C + 1
    O(1, C) = V
Next
AySqH = O
End Function

Function AySqV(Ay) As Variant()
If AyIsEmp(Ay) Then Exit Function
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
If AyIsEmp(Ay) Then AySrt = Ay: Exit Function
Dim Ix&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    O = AyIns(O, Ay(J), AySrt__Ix(O, Ay(J), Des))
Next
AySrt = O
End Function

Function AySrtInToIxAy(Ay, Optional Des As Boolean) As Long()
If AyIsEmp(Ay) Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = AyIns(O, J, AySrtInToIxAy__Ix(O, Ay, Ay(J), Des))
Next
AySrtInToIxAy = O
End Function

Function AyTrim(A) As String()
If AyIsEmp(A) Then Exit Function
Dim U&
    U = UB(A)
Dim O$()
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = Trim(A(J))
    Next
AyTrim = O
End Function

Function AyUniq(Ay)
Dim O: O = Ay: Erase O
Dim V
If Not AyIsEmp(Ay) Then
    For Each V In Ay
        PushNoDup O, V
    Next
End If
AyUniq = O
End Function

Function AyVSq(Ay)
Dim O
Dim N&
N = Sz(Ay)
ReDim O(1 To N, 1 To 1)
Dim J&
For J = 1 To N
    O(J, 1) = Ay(J - 1)
Next
AyVSq = O
End Function

Function AyWdt%(Ay)
If AyIsEmp(Ay) Then Exit Function
Dim O%, I
For Each I In Ay
    O = Max(O, Len(I))
Next
AyWdt = O
End Function

Function AyWh(Ay, Fmix&, Toix&)
Dim O: O = Ay: Erase O
AyWh = O
If AyIsEmp(Ay) Then Exit Function
If Fmix < 0 Then Exit Function
If Toix < 0 Then Exit Function
Dim J&
For J = Fmix To Toix
    Push O, Ay(J)
Next
AyWh = O
End Function

Function AyWhDist(Ay)
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    PushNoDup O, I
Next
AyWhDist = O
End Function

Function AyWhDup(Ay)
Dim O: O = Ay: Erase O
Dim GpDry(): GpDry = AyGpDry(Ay)
If AyIsEmp(GpDry) Then AyWhDup = O: Exit Function
Dim Dr
For Each Dr In GpDry
    If Dr(1) > 1 Then Push O, Dr(0)
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

Function AyWhExclIxAy(Ay, IxAy)
'IxAy holds index if Ay to be remove.  It has been sorted else will be stop
Ass AyIsSrt(Ay)
Ass AyIsSrt(IxAy)
Dim J&
Dim O: O = Ay
For J = UB(IxAy) To 0 Step -1
    O = AyRmvEleAt(O, CLng(IxAy(J)))
Next
AyWhExclIxAy = O
End Function

Function AyWhFm(Ay, Fmix&)
Dim O: O = Ay: Erase O
If 0 <= Fmix And Fmix <= UB(Ay) Then
    Dim J&
    For J = Fmix To UB(Ay)
        Push O, Ay(J)
    Next
End If
AyWhFm = O
End Function

Function AyWhFstNEle(Ay, N&)
Dim O: O = Ay
ReDim Preserve O(N - 1)
AyWhFstNEle = O
End Function

Function AyWhIxAy(Ay, IxAy, Optional CrtEmpEle_IfReqEleNotFound As Boolean)
'Return a subset of {Ay} by {IxAy}
Ass IsArray(Ay)
Ass IsArray(IxAy)
Dim O
    O = Ay: Erase O
    If AyIsEmp(IxAy) Then
        Er "AyWhIxAy", "Given [IxAy] is empty.  [IxAy] is IxAy of subset of {Ay}", IxAy
        GoTo X
    End If
    Dim U&
    U = UB(IxAy)
    Dim J&, Ix
    ReDim O(U)
    For J = 0 To U
        Ix = IxAy(J)
        If Ix = -1 Then
            If Not CrtEmpEle_IfReqEleNotFound Then
                Er "AyWhIxAy", "Given {IxAy} contains -1", IxAy
            End If
        Else
            If IsObject(Ay(Ix)) Then
                Set O(J) = Ay(Ix)
            Else
                O(J) = Ay(Ix)
            End If
        End If
    Next
X:
AyWhIxAy = O
End Function

Function AyWhLik(Ay, Lik$) As String()
If AyIsEmp(Ay) Then Exit Function
Dim O$()
Dim I
For Each I In Ay
    If I Like Lik Then Push O, I
Next
AyWhLik = O
End Function

Function AyWhLikAy(Ay, LikAy$()) As String()
If AyIsEmp(Ay) Then Exit Function
If AyIsEmp(LikAy) Then Exit Function
Dim I, Lik, O$()
For Each I In Ay
    For Each Lik In LikAy
        If I Like Lik Then
            Push O, I
            Exit For
        End If
    Next
Next
AyWhLikAy = O
End Function

Function AyWhMulEle(Ay)
'Return Set of Element as array in {Ay} having 2 or more element
Dim Dry(): Dry = AyGpDry(Ay)
Dim O: O = Ay: Erase O
Dim Dr
If Not AyIsEmp(Dry) Then
    For Each Dr In Dry
        If Dr(1) > 1 Then
            Push O, Dr(0)
        End If
    Next
End If
AyWhMulEle = O
End Function

Function AyWhPatn(Ay, Patn$) As String()
If AyIsEmp(Ay) Then Exit Function
Dim I, O$()
Dim R As RegExp
Set R = Re(Patn)
For Each I In Ay
    If R.Test(I) Then Push O, I
Next
AyWhPatn = O
End Function

Function AyWhPatnIx(Ay, Patn$) As Long()
If AyIsEmp(Ay) Then Exit Function
Dim I, O&(), J&
Dim R As RegExp
Set R = Re(Patn)
For Each I In Ay
    If R.Test(I) Then Push O, J
    J = J + 1
Next
AyWhPatnIx = O
End Function

Function AyWhPfx(Ay, Pfx$) As String()
If AyIsEmp(Ay) Then Exit Function
Dim O$()
Dim I
For Each I In Ay
    If HasPfx(I, Pfx) Then Push O, I
Next
AyWhPfx = O
End Function

Function AyWhPred(Ay, PredMthNm$, ParamArray Ap())
Dim O: O = Ay: Erase O
Dim I
Dim Av()
    Av = Ap
    Av = AyIns(Av)
For Each I In Ay
    Asg I, Av(0)
    If RunAv(PredMthNm, Av) Then
        Push O, I
    End If
Next
AyWhPred = O
End Function

Function AyWhSfx(Ay, Sfx$) As String()
If AyIsEmp(Ay) Then Exit Function
Dim O$()
Dim I
For Each I In Ay
    If HasSfx(CStr(I), Sfx) Then Push O, I
Next
AyWhSfx = O
End Function

Function AyWhSngEle(Ay)
'Return Set of Element as array in {Ay} having 2 or more element
Dim Dry(): Dry = AyGpDry(Ay)
Dim O: O = Ay: Erase O
Dim Dr
If Not AyIsEmp(Dry) Then
    For Each Dr In Dry
        If Dr(1) = 1 Then
            Push O, Dr(0)
        End If
    Next
End If
AyWhSngEle = O
End Function

Function AyWs(Ay, Optional WsNm$, Optional Vis As Boolean) As Worksheet
Stop
'Dim O As Worksheet: Set O = NewWs(WsNm, Vis)
'SqRg AyVSq(Ay), WsA1(O)
'Set AyWs = O
End Function

Function AyZip(A1, A2) As Variant()
Dim U1&: U1 = UB(A1)
Dim U2&: U2 = UB(A2)
Dim U&: U = Max(U1, U2)
Dim O(), J&
ReSz O, U
For J = 0 To U
    If U1 >= J Then
        If U2 >= J Then
            O(J) = Array(A1(J), A2(J))
        Else
            O(J) = Array(A1(J), Empty)
        End If
    Else
        If U2 >= J Then
            O(J) = Array(, A2(J))
        Else
            Stop
        End If
    End If
Next
AyZip = O
End Function

Function AyZipAp(A1, ParamArray Ap()) As Variant()
Dim Av(): Av = Ap
Dim UCol%
    UCol = UB(Av)

Dim URow1&
    URow1 = UB(A1)

Dim URow&
Dim URowAy&()
    Dim J%, IURow%
    URow = URow1
    For J = 0 To UB(Av)
        IURow = UB(Av(J))
        Push URowAy, IURow
        If IURow > URow Then URow = IURow
    Next

Dim ODry()
    Dim Dr()
    ReSz ODry, URow
    Dim I%
    For J = 0 To URow
        Erase Dr
        If URow1 >= J Then
            Push Dr, A1(J)
        Else
            Push Dr, Empty
        End If
        For I = 0 To UB(Av)
            If URowAy(I) >= J Then
                Push Dr, Av(I)(J)
            Else
                Push Dr, Empty
            End If
        Next
        ODry(J) = Dr
    Next
AyZipAp = ODry
End Function

Sub AyBrw(Ay, Optional Fnn$)
Dim T$
T = TmpFt("AyBrw", Fnn)
AyWrt Ay, T
FtBrw T
End Sub

Sub AyCastEle(Ay, ParamArray OAp())
Dim V, J%
For Each V In Ay
    If Not IsMissing(OAp(J)) Then
        Asg Ay(J), OAp(J)
    End If
    J = J + 1
Next
End Sub

Sub AyChkEq(Ay1, Ay2, Optional Nm1$ = "Ay1", Optional Nm2$ = "Ay2")
Chk AyEqChk(Ay1, Ay2, Nm1, Nm2)
End Sub

Sub AyDmp(Ay, Optional WithIx As Boolean)
If AyIsEmp(Ay) Then Exit Sub
Dim I
If WithIx Then
    Dim J&
    For Each I In Ay
        Debug.Print J; ": "; I
        J = J + 1
    Next
Else
    For Each I In Ay
        Debug.Print I
    Next
End If

End Sub
Sub ZZ_AyAsgAp()
Dim O%, A$
AyAsgAp Array(234, "abc"), O, A
Ass O = 234
Ass A = "abc"
End Sub
Sub AyAsgAp(A, ParamArray OAp())
Dim Av(): Av = OAp
Dim J&
For J = 0 To UB(Av)
    Asg A(J), OAp(J)
Next
End Sub
Sub AyIxAyAsgAp(A, IxAy&(), ParamArray OAp())
Dim J&
For J = 0 To UB(IxAy)
    Asg A(IxAy(J)), OAp(J)
Next
End Sub

Sub AyRmvLasNEle(Ay, Optional NEle% = 1)
ReDim Preserve Ay(UB(Ay) - NEle)
End Sub

Sub AyWrt(Ay, Ft)
StrWrt JnCrLf(Ay), Ft
End Sub

Private Sub AyGpDry__Upd(OGpDry(), Itm)
Dim J&
For J = 0 To UB(OGpDry)
    If OGpDry(J)(0) = Itm Then
        OGpDry(J)(1) = OGpDry(J)(1) + 1
        Exit Sub
    End If
Next
Push OGpDry, Array(Itm, 1)
End Sub

Private Sub AyIxAy__ChkNotFound(IxAy&(), A, SubAy)
If IxAy_IsAllGE0(IxAy) Then Exit Sub
Dim J&, SomEle(), SomEleIx&()
For J = 0 To UB(IxAy)
    If IxAy(J) = -1 Then
        Push SomEleIx, J
        Push SomEle, SubAy(J)
    End If
Next
Er "AyIxAy__ChkNotFound", "{SomEle} with {Ix} in {SubAy} are not found in Given {Ay}", SomEle, SomEleIx, SubAy, A, SomEleIx
End Sub

Private Function AySrt__Ix&(Ay, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ay
        If V > I Then AySrt__Ix = O: Exit Function
        O = O + 1
    Next
    AySrt__Ix = O
    Exit Function
End If
For Each I In Ay
    If V < I Then AySrt__Ix = O: Exit Function
    O = O + 1
Next
AySrt__Ix = O
End Function

Sub ZZZ__Tst()
ZZ_AyTrim
End Sub

Private Function AySrtInToIxAy__Ix&(Ix&(), A, V, Des As Boolean)
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

Private Sub ZZ_AyAdd()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyAdd(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
AyChkEq Exp, Act
AyChkEq Ay1, Array(1, 2, 2, 2, 4, 5)
AyChkEq Ay2, Array(2, 2)
End Sub

Private Sub ZZ_AyEqChk()
AyDmp AyEqChk(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Private Sub ZZZ_AyBrkInto3Ay()
Dim A(): A = Array(1, 2, 3, 4)
Dim Act(): Act = AyBrkInto3Ay(A, 1, 2)
Ass Sz(Act) = 3
Ass AyIsEq(Act(0), Array(1))
Ass AyIsEq(Act(1), Array(2, 3))
Ass AyIsEq(Act(2), Array(4))
End Sub

Private Sub ZZ_AyGpDry()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = AyGpDry(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
Stop
'AssEqDry Act, Exp
End Sub

Private Sub ZZ_AyHasDupEle()
Ass AyHasDupEle(Array(1, 2, 3, 4)) = False
Ass AyHasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Private Sub ZZ_AyIntAy()
Dim Act%(): Act = AyIntAy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub ZZ_AyMap()
Dim Act: Act = AyMap(Array(1, 2, 3, 4), "Mul2")
Ass Sz(Act) = 4
Ass Act(0) = 2
Ass Act(1) = 4
Ass Act(2) = 6
Ass Act(3) = 8
End Sub

Private Sub ZZ_AyMapSy()
Dim Ay$(): Ay = AyMapSy(Array("skldfjdf", "aa"), "RmvFstChr")
Stop
End Sub

Private Sub ZZ_AyMinus()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
AyChkEq Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
AyChkEq Exp, Act
End Sub

Private Sub ZZ_AyRmvEmpEleAtEnd()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyRmvEmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub ZZ_AyRmvFmTo()
Dim A
Dim FmTo As FmTo
Dim Act
A = SplitSpc("a b c d e")
FmTo.Fmix = 1
FmTo.Toix = 2
Act = AyRmvFmTo(A, FmTo)
Ass Sz(Act) = 3
Ass JnSpc(Act) = "a d e"
End Sub

Private Sub ZZ_AySrt()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                    Act = AySrt(A):        AyChkEq Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(A, True): AyChkEq Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(A):       AyChkEq Exp, Act
'-----------------
Erase A
Push A, ":PjUpdTm:Sub"
Push A, ":MthBrk:Function"
Push A, "~~:Tst:Sub"
Push A, ":PjTmNy_WithEr:Function"
Push A, "~Private:JnContinueLin__Tst:Sub"
Push A, "Private:HasPfx:Function"
Push A, "Private:MdMthDrs_FunBdyLy:Function"
Push A, "Private:SrcMthLx_ToLx:Function"
Erase Exp
Push Exp, ":PjTmNy_WithEr:Function"
Push Exp, ":PjUpdTm:Sub"
Push Exp, ":MthBrk:Function"
Push Exp, "Private:HasPfx:Function"
Push Exp, "Private:MdMthDrs_FunBdyLy:Function"
Push Exp, "Private:SrcMthLx_ToLx:Function"
Push Exp, "~Private:JnContinueLin__Tst:Sub"
Push Exp, "~~:Tst:Sub"
Act = AySrt(A)
AyChkEq Exp, Act
End Sub

Private Sub ZZ_AySrtInToIxAy()
Dim A: A = Array("A", "B", "C", "D", "E")
AyChkEq Array(0, 1, 2, 3, 4), AySrtInToIxAy(A)
AyChkEq Array(4, 3, 2, 1, 0), AySrtInToIxAy(A, True)
End Sub

Private Sub ZZ_AySy()
Dim Act$(): Act = AySy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub ZZ_AyTrim()
AyDmp AyTrim(Array(1, 2, 3, "  a"))
End Sub

Private Sub ZZ_AyWhExclAtCnt()
Dim A(): A = Array(1, 2, 3, 4, 5)
Dim Act: Act = AyWhExclAtCnt(A, 1, 2)
AyChkEq Array(1, 4, 5), Act
End Sub

Private Sub ZZ_AyWhExclIxAy()
Dim A(): A = Array("a", "b", "c", "d", "e", "f")
Dim IxAy: IxAy = Array(1, 3)
Dim Exp: Exp = Array("a", "c", "e", "f")
Dim Act: Act = AyWhExclIxAy(A, IxAy)
Ass Sz(Act) = 4
Dim J%
For J = 0 To 3
    Ass Act(J) = Exp(J)
Next
End Sub

