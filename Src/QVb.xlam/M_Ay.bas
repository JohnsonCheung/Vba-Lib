Attribute VB_Name = "M_Ay"
Option Explicit

Property Get AyAdd(Ay1, Ay2)
Dim O: O = Ay1
PushAy O, Ay2
AyAdd = O
End Property
Property Get AyFstNEle(A, N&)
Dim O: O = A
ReDim Preserve O(N - 1)
AyFstNEle = O
End Property
Property Get AyAddAp(Ay, ParamArray Itm_or_Ay_Ap())
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
End Property

Property Get AyAddPfx(Ay, Pfx) As String()
If AyIsEmp(Ay) Then Exit Property
Dim O$(), J&, U&
U = UB(Ay)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Ay(J)
Next
AyAddPfx = O
End Property

Property Get AyAddPfxSfx(Ay, Pfx, Sfx) As String()
Dim O$(), J&, U&
If AyIsEmp(Ay) Then Exit Property
U = UB(Ay)
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & Ay(J) & Sfx
Next
AyAddPfxSfx = O
End Property

Property Get AyAddSfx(Ay, Sfx) As String()
Dim O$(), J&, U&
If AyIsEmp(Ay) Then Exit Property
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Ay(J) & Sfx
Next
AyAddSfx = O
End Property

Property Get AyAlignL(Ay) As String()
If AyIsEmp(Ay) Then Exit Property
Dim W%: W = AyWdt(Ay)
Dim O$(), I
For Each I In Ay
    Push O, AlignL(I, W)
Next
AyAlignL = O
End Property

Property Get AyBrk3ByIx(Ay, FmIx&, ToIx&)
AyBrk3ByIx = AyFmTo_Brk(Ay, FmTo(FmIx, ToIx))
End Property

Property Get AyCellSy(Ay, Optional ShwZer As Boolean) As String()
Dim O$(), I, J&, U&
U = UB(Ay)
ReSz O, U
For Each I In Ay
    O(J) = ToCellStr(I)
    J = J + 1
Next
End Property

Property Get AyConst1_Dry(Ay, C) As Variant()
'C1Dry is Dry with 2 column and Col1 is const
Dim U&, J&
U = UB(Ay)
Dim O()
ReSz O, U
For J = 0 To U
    O(J) = Array(C, Ay(J))
Next
AyConst1_Dry = O
End Property

Property Get AyConst2_Dry(Ay, C) As Variant()
'C2Dry is Dry with 2 column and Col1 is const
Dim U&, J&
U = UB(Ay)
Dim O()
ReSz O, U
For J = 0 To U
    O(J) = Array(Ay(J), C)
Next
AyConst2_Dry = O
End Property

Property Get AyDblAy(Ay) As Double()
AyDblAy = AyCast(Ay, EmpDblAy)
End Property

Property Get AyDic(Ay, Optional V = True) As Dictionary
Dim O As New Dictionary, I
If Not AyIsEmp(Ay) Then
    For Each I In Ay
        O.Add I, V
    Next
End If
Set AyDic = O
End Property

Property Get AyDry(Ay) As Variant()
Dim O(), J&
Dim U&: U = UB(Ay)
ReSz O, U
For J = 0 To U
    O(J) = Array(Ay(J))
Next
AyDry = O
End Property

Property Get AyDupAy(Ay)
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
End Property

Property Get AyEqChk(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act") As String()
Dim U&: U = UB(Ay1)
Dim O$()
    If U <> UB(Ay2) Then Push O, FmtQQ("Array [?] and [?] has different Sz: [?] [?]", Ay1Nm, Ay2Nm, Sz(Ay1), Sz(Ay2)): GoTo X
If AyIsEmp(Ay1) Then Exit Property
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
If M_Is.IsEmp(O1) Then Exit Property
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
End Property

Property Get AyFmTo_Brk(Ay, B As FmTo) As Variant()
Ass FmTo_HasU(B, UB(Ay))
Dim O(2)
O(0) = AyWhFmTo(Ay, FmTo(0, B.FmIx - 1))
O(1) = AyWhFmTo(Ay, B)
O(2) = AyWhFmTo(Ay, FmTo(B.FmIx + 1, UB(Ay)))
AyFmTo_Brk = O
End Property

Property Get AyGpDry(Ay) As Variant()
If AyIsEmp(Ay) Then Exit Property
Dim O(), I
For Each I In Ay
    AyGpDry__Upd O, I
Next
AyGpDry = O
End Property

Property Get AyHas(Ay, Itm) As Boolean
If AyIsEmp(Itm) Then Exit Property
Dim I
For Each I In Ay
    If I = Itm Then AyHas = True: Exit Property
Next
End Property

Property Get AyHasDupEle(Ay) As Boolean
If AyIsEmp(Ay) Then Exit Property
Dim Pool: Pool = Ay: Erase Pool
Dim I
For Each I In Ay
    If AyHas(Pool, I) Then AyHasDupEle = True: Exit Property
    Push Pool, I
Next
End Property

Property Get AyHasNegOne(Ay) As Boolean
Dim V
If AyIsEmp(Ay) Then Exit Property
For Each V In Ay
    If V = -1 Then AyHasNegOne = True: Exit Property
Next
End Property

Property Get AyHasSubAy(Ay, SubAy) As Boolean
If AyIsEmp(Ay) Then Exit Property
If AyIsEmp(SubAy) Then PmEr
Dim I
For Each I In SubAy
    If Not AyHas(Ay, I) Then Exit Property
Next
End Property

Property Get AyIncNForEachEle(Ay, Optional N& = 1)
Dim O: O = Ay
Dim J&
For J = 0 To UB(Ay)
    O(J) = O(J) + N
Next
AyIncNForEachEle = O
End Property

Property Get AyIns(Ay, Optional Ele, Optional At&)
Const CSub$ = "AyIns"
Dim N&: N = Sz(Ay)
If 0 > At Or At > N Then
    Stop
    Er CSub, "{At} is outside {Ay-UB}", At, N - 1
End If
Dim O
    O = Ay
    ReDim Preserve O(N)
    Dim J&
    For J = N To At + 1 Step -1
        Asg O(J - 1), O(J)
    Next
    O(At) = Ele
AyIns = O
End Property

Property Get AyIntersect(Ay1, Ay2)
Dim O: O = Ay1: Erase O
If AyIsEmp(Ay1) Then GoTo X
If AyIsEmp(Ay2) Then GoTo X
Dim V
For Each V In Ay1
    If AyHas(Ay2, V) Then O.Push V
Next
X:
AyIntersect = O
End Property

Property Get AyIsAllEleHasPfx(A, Pfx$) As Boolean
If AyIsEmp(A) Then Exit Property
Dim I
For Each I In A
   If Not HasPfx(I, Pfx) Then Exit Property
Next
AyIsAllEleHasPfx = True
End Property

Property Get AyIsAllEleHasVal(Ay) As Boolean
If AyIsEmp(Ay) Then Exit Property
Dim I
For Each I In Ay
    If M_Is.IsEmp(I) Then Exit Property
Next
AyIsAllEleHasVal = True
End Property

Property Get AyIsAllEq(Ay) As Boolean
If AyIsEmp(Ay) Then AyIsAllEq = True: Exit Property
Dim T: T = Ay(0)
Dim J&
For J = 1 To UB(Ay)
    If Ay(J) = T Then Exit Property
Next
AyIsAllEq = True
End Property

Property Get AyIsAllStr(Ay) As Boolean
If Sz(Ay) = 0 Then Exit Property
Dim K
For Each K In Ay
    If Not IsStr(K) Then Exit Property
Next
AyIsAllStr = True
End Property

Property Get AyIsEmp(V) As Boolean
AyIsEmp = Sz(V) = 0
End Property

Property Get AyIsEq(Ay1, Ay2) As Boolean
Dim U&: U = UB(Ay1): If U <> UB(Ay2) Then Exit Property
Dim J&
For J = 0 To U
   If Ay1(J) <> Ay2(J) Then Exit Property
Next
AyIsEq = True
End Property

Property Get AyIsEqSz(Ay, B) As Boolean
AyIsEqSz = Sz(Ay) = Sz(B)
End Property

Property Get AyIsSamSz(Ay1, Ay2) As Boolean
AyIsSamSz = Sz(Ay1) = Sz(Ay2)
End Property

Property Get AyIsSrt(Ay) As Boolean
Dim J&
For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Property
Next
AyIsSrt = True
End Property

Property Get AyIx&(Ay, Itm)
Dim J&
For J = 0 To UB(Ay)
    If Ay(J) = Itm Then AyIx = J: Exit Property
Next
AyIx = -1
End Property

Property Get AyIxAy(Ay, SubAy, Optional ChkNotFound As Boolean, Optional SkipNotFound As Boolean) As Long()
If AyIsEmp(SubAy) Then Exit Property
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
End Property

Property Get AyLasEle(Ay)
AyLasEle = Ay(UB(Ay))
End Property

Property Get AyMap(Ay, MthNm$, ParamArray Ap()) As Variant()
If AyIsEmp(Ay) Then Exit Property
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
End Property

Property Get AyMapAsgAy(Ay, OAy, MthNm$, ParamArray Ap())
If AyIsEmp(Ay) Then Exit Property
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
End Property

Property Get AyMapAsgSy(Ay, MthNm$, ParamArray Ap()) As String()
If AyIsEmp(Ay) Then Exit Property
Dim Av(): Av = Ap
If AyIsEmp(Av) Then
    AyMapAsgSy = AyMap_Sy(Ay, MthNm)
    Exit Property
End If
Dim I, J&
Dim O$()
    ReDim O(UB(Ay))
    Av = AyIns(Av)
    For Each I In Ay
        Asg I, Av(0)
        Asg RunAv(MthNm, Av), O(J)
        J = J + 1
    Next
AyMapAsgSy = O
End Property

Property Get AyMapInto(Ay, Obj, GetNm$, OIntoAy)
Dim O: O = OIntoAy: Erase O
Dim J&, U&
Dim Arg
U = UB(Ay)
ReSz O, U
For J = 0 To U
    Asg Ay(J), Arg
    Asg CallByName(Obj, GetNm, VbGet, Arg), O(J)
Next
AyMapInto = O
End Property

Property Get AyMap_Lng(Ay, MapMthNm$) As Long()
AyMap_Lng = AyLngAy(AyMap(Ay, MapMthNm))
End Property

Property Get AyMap_Sy(Ay, MapMthNm$) As String()
AyMap_Sy = AySy(AyMap(Ay, MapMthNm))
End Property

Property Get AyMax(Ay)
If AyIsEmp(Ay) Then Exit Property
Dim O, I
For Each I In Ay
    If I > O Then O = I
Next
AyMax = O
End Property

Property Get AyMinus(Ay1, Ay2)
If AyIsEmp(Ay1) Then AyMinus = Ay1: Exit Property
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
End Property

Property Get AyMinusAp(Ay, ParamArray AyAp())
Dim O: O = Ay
Dim Av(): Av = AyAp
Dim Ay1, V
For Each Ay1 In Av
    If AyIsEmp(O) Then GoTo X
    O = AyMinus(O, Ay1)
Next
X:
AyMinusAp = O
End Property

Property Get AyNoDupAy(Ay)
Dim O: O = Ay
Erase O
Dim I
If AyIsEmp(Ay) Then AyNoDupAy = O: Exit Property
For Each I In Ay
    PushNoDup O, I
Next
AyNoDupAy = O
End Property

Property Get AyRTrim(Ay) As String()
If AyIsEmp(Ay) Then Exit Property
Dim O$(), I
For Each I In Ay
    M_Ay.Push O, RTrim(I)
Next
AyRTrim = O
End Property

Property Get AyReOrd(Ay, PartialIxAy&())
Dim I&()
    I = PartialIxAy_CompleteIxAy(PartialIxAy, UB(Ay))
Dim O
    O = Ay: Erase O
    Dim J&
    For J = 0 To UB(I)
        Push O, Ay(I(J))
    Next
AyReOrd = O
End Property

Property Get AyRmvEle(Ay, Ele)
Dim Ix&: Ix = AyIx(Ay, Ele): If Ix = -1 Then AyRmvEle = Ay: Exit Property
AyRmvEle = AyRmvEleAt(Ay, AyIx(Ay, Ele))
End Property

Property Get AyRmvEleAt(Ay, Optional At&)
AyRmvEleAt = AyWhExclAtCnt(Ay, At)
End Property

Property Get AyRmvEmpEle(Ay)
If AyIsEmp(Ay) Then AyRmvEmpEle = Ay: Exit Property
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    If Not IsEmp(I) Then Push O, I
Next
AyRmvEmpEle = O
End Property

Property Get AyRmvEmpEleAtEnd(Ay)
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
End Property

Property Get AyRmvFmTo(Ay, FmTo As FmTo)
Dim O
    O = Ay
    If Not FmTo_IsVdt(FmTo) Or AyIsEmp(Ay) Then
        Dim FmI&, ToI&
        FmI = FmTo.FmIx
        ToI = FmTo.ToIx
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
End Property

Property Get AyRmvFstEle(Ay)
AyRmvFstEle = AyRmvEleAt(Ay)
End Property

Property Get AyRmvLasChr(A) As String()
Dim O$(), I
For Each I In A
    Push O, RmvLasChr(I)
Next
AyRmvLasChr = O
End Property

Property Get AyRmvLasEle(Ay)
AyRmvLasEle = AyRmvEleAt(Ay, UB(Ay))
End Property

Property Get AyRmvPfx(Ay, Pfx) As String()
If AyIsEmp(Ay) Then Exit Property
Dim U&: U = UB(Ay)
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = RmvPfx(Ay(J), Pfx)
Next
AyRmvPfx = O
End Property

Property Get AyRpl(Ay, FmTo As FmTo, AySeg)
Dim A()
    A = AyFmTo_Brk(Ay, FmTo)
Dim O
    O = Ay(0): Erase O
    PushAy O, AySeg
    PushAy O, Ay(2)
AyRpl = O
End Property

Property Get AyShift(OAy)
AyShift = OAy(0)
OAy = AyRmvFstEle(OAy)
End Property

Property Get AySqH(Ay) As Variant()
If AyIsEmp(Ay) Then Exit Property
Dim O(), C%
ReDim O(1 To 1, 1 To Sz(Ay))
C = 0
Dim V
For Each V In Ay
    C = C + 1
    O(1, C) = V
Next
AySqH = O
End Property

Property Get AySqV(Ay) As Variant()
If AyIsEmp(Ay) Then Exit Property
Dim O(), R&
ReDim O(1 To Sz(Ay), 1 To 1)
R = 0
Dim V
For Each V In Ay
    R = R + 1
    O(R, 1) = V
Next
AySqV = O
End Property

Property Get AySrt(Ay, Optional Des As Boolean)
If AyIsEmp(Ay) Then AySrt = Ay: Exit Property
Dim Ix&, V, J&
Dim O: O = Ay: Erase O
Push O, Ay(0)
For J = 1 To UB(Ay)
    O = AyIns(O, Ay(J), AySrt__Ix(O, Ay(J), Des))
Next
AySrt = O
End Property

Property Get AySrtInToIxAy(Ay, Optional Des As Boolean) As Long()
If AyIsEmp(Ay) Then Exit Property
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(Ay)
    O = AyIns(O, J, AySrtInToIxAy_Ix(O, Ay, Ay(J), Des))
Next
AySrtInToIxAy = O
End Property

Property Get AyTrim(A) As String()
If AyIsEmp(A) Then Exit Property
Dim U&
    U = UB(A)
Dim O$()
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = Trim(A(J))
    Next
AyTrim = O
End Property

Property Get AyUniq(Ay)
Dim O: O = Ay: Erase O
Dim V
If Not AyIsEmp(Ay) Then
    For Each V In Ay
        PushNoDup O, V
    Next
End If
AyUniq = O
End Property

Property Get AyVSq(Ay)
Dim O
Dim N&
N = Sz(Ay)
ReDim O(1 To N, 1 To 1)
Dim J&
For J = 1 To N
    O(J, 1) = Ay(J - 1)
Next
AyVSq = O
End Property

Property Get AyWdt%(Ay)
If AyIsEmp(Ay) Then Exit Property
Dim O%, I
For Each I In Ay
    O = Max(O, Len(I))
Next
AyWdt = O
End Property

Property Get AyWh(Ay, FmIx&, ToIx&)
Dim O: O = Ay: Erase O
AyWh = O
If AyIsEmp(Ay) Then Exit Property
If FmIx < 0 Then Exit Property
If ToIx < 0 Then Exit Property
Dim J&
For J = FmIx To ToIx
    Push O, Ay(J)
Next
AyWh = O
End Property

Property Get AyWhDist(Ay)
Dim O: O = Ay: Erase O
Dim I
For Each I In Ay
    PushNoDup O, I
Next
AyWhDist = O
End Property

Property Get AyWhDup(Ay)
Dim O: O = Ay: Erase O
Dim GpDry(): GpDry = AyGpDry(Ay)
If AyIsEmp(GpDry) Then AyWhDup = O: Exit Property
Dim Dr
For Each Dr In GpDry
    If Dr(1) > 1 Then Push O, Dr(0)
Next
AyWhDup = O
End Property

Property Get AyWhExclAtCnt(Ay, At&, Optional Cnt& = 1)
If Cnt <= 0 Then AyWhExclAtCnt = Ay: Exit Property
Dim U&: U = UB(Ay)
If At > U Then Stop
If At < 0 Then Stop
If U = 0 Then AyWhExclAtCnt = Ay: Exit Property
Dim O: O = Ay
Dim J&
For J = At To U - Cnt
    O(J) = O(J + Cnt)
Next
ReDim Preserve O(U - Cnt)
AyWhExclAtCnt = O
End Property

Property Get AyWhExclIxAy(Ay, IxAy)
'IxAy holds index if Ay to be remove.  It has been sorted else will be stop
Ass AyIsSrt(Ay)
Ass AyIsSrt(IxAy)
Dim J&
Dim O: O = Ay
For J = UB(IxAy) To 0 Step -1
    O = AyRmvEleAt(O, CLng(IxAy(J)))
Next
AyWhExclIxAy = O
End Property

Property Get AyWhFm(Ay, FmIx&)
Dim O: O = Ay: Erase O
If 0 <= FmIx And FmIx <= UB(Ay) Then
    Dim J&
    For J = FmIx To UB(Ay)
        Push O, Ay(J)
    Next
End If
AyWhFm = O
End Property

Property Get AyWhFmTo(Ay, FmTo As FmTo)
AyWhFmTo = AyWh(Ay, FmTo.FmIx, FmTo.ToIx)
End Property

Property Get AyWhFstNEle(Ay, N&)
Dim O: O = Ay
ReDim Preserve O(N - 1)
AyWhFstNEle = O
End Property

Property Get AyWhIxAy(Ay, IxAy, Optional CrtEmpEle_IfReqEleNotFound As Boolean)
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
End Property

Property Get AyWhLik(Ay, Lik$) As String()
If AyIsEmp(Ay) Then Exit Property
Dim O$()
Dim I
For Each I In Ay
    If I Like Lik Then Push O, I
Next
AyWhLik = O
End Property

Property Get AyWhLikAy(Ay, LikAy$()) As String()
If AyIsEmp(Ay) Then Exit Property
If AyIsEmp(LikAy) Then Exit Property
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
End Property

Property Get AyWhMulEle(Ay)
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
End Property

Property Get AyWhPatnIx(Ay, Patn$) As Long()
If AyIsEmp(Ay) Then Exit Property
Dim I, O&(), J&
Dim R As RegExp
Set R = Re(Patn)
For Each I In Ay
    If ReTst(R, I) Then Push O, J
    J = J + 1
Next
AyWhPatnIx = O
End Property

Property Get AyWhSfx(Ay, Sfx$) As String()
If AyIsEmp(Ay) Then Exit Property
Dim O$()
Dim I
For Each I In Ay
    If HasSfx(CStr(I), Sfx) Then Push O, I
Next
AyWhSfx = O
End Property

Property Get AyWhSngEle(Ay)
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
End Property

Property Get AyWh_ByIxAy(Ay, IxAy)
Dim O: O = Ay: Erase O
Dim J%
For J = 0 To UB(IxAy)
    Push O, Ay(IxAy(J))
Next
AyWh_ByIxAy = O
End Property

Property Get AyWh_ByMth(Ay, WhMthNm$, ParamArray Ap())
Dim O: O = Ay: Erase O
Dim I
Dim Av()
    Av = Ap
    Av = AyIns(Av)
For Each I In Ay
    Asg I, Av(0)
    If RunAv(WhMthNm, Av) Then
        Push O, I
    End If
Next
AyWh_ByMth = O
End Property

Property Get AyWh_ByPatn(Ay, Patn$) As String()
If AyIsEmp(Ay) Then Exit Property
Dim I, O$()
Dim R As RegExp
Set R = Re(Patn)
For Each I In Ay
    If ReTst(R, I) Then Push O, I
Next
AyWh_ByPatn = O
End Property

Property Get AyWh_ByPfx(Ay, Pfx$) As String()
If AyIsEmp(Ay) Then Exit Property
Dim O$()
Dim I
For Each I In Ay
    If HasPfx(CStr(I), Pfx) Then Push O, I
Next
AyWh_ByPfx = O
End Property

Property Get AyWs(Ay, Optional WsNm$, Optional Vis As Boolean) As Worksheet
Stop
'Dim O As Worksheet: Set O = NewWs(WsNm, Vis)
'SqRg AyVSq(Ay), WsA1(O)
'Set AyWs = O
End Property

Property Get AyZip(A1, A2) As Variant()
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
End Property

Property Get AyZipAp(A1, ParamArray Ap()) As Variant()
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
End Property

Property Get FnyIxAy(A$(), SubFny0) As Integer()
Dim SubFny$(): SubFny = DftNy(SubFny0)
If AyIsEmp(SubFny) Then Stop
Dim O%(), U&, J%
U = UB(SubFny)
ReSz O, U
For J = 0 To U
    O(J) = AyIx(A, SubFny(J))
    If O(J) = -1 Then Stop
Next
End Property

Property Get IxAy_IsAllGE0(IxAy&()) As Boolean
Dim J&
For J = 0 To UB(IxAy)
    If IxAy(J) = -1 Then Exit Property
Next
IxAy_IsAllGE0 = True
End Property

Property Get IxAy_IsParitial_of_0toU(IxAy, U&) As Boolean
Const CSub$ = "Ass IxAy_IsParitial_of_0toU"
Const Msg$ = "{IxAy} is not PartialIx-of-{U}." & _
"|PartialIxAy-Of-U is defined as:" & _
"|It should be Lng()" & _
"|It should have 0 to U elements" & _
"|It should have each element of value between 0 and U" & _
"|It should have no dup element" & _
"|All elements should have value equal or less than U"

If Not IsLngAy(IxAy) Then Exit Property
If AyIsEmp(IxAy) Then IxAy_IsParitial_of_0toU = True: Exit Property
If AyHasDupEle(IxAy) Then Exit Property
Dim I
For Each I In IxAy
   If 0 > I Or I > U Then Exit Property
Next
IxAy_IsParitial_of_0toU = True
End Property

Property Get PartialIxAy_CompleteIxAy(PartialIxAy&(), U&) As Long()
'Des:Make a complete-IxAy-of-U by partialIxAy
'Des:A complete-IxAy-Of-U is defined as
'Des:it has (U+1)-elements,
'Des:it does not have dup element
'Des:it has all element of value between 0 and U
Ass IxAy_IsParitial_of_0toU(PartialIxAy, U)
Dim I&(): I = ULngSeq(U)
PartialIxAy_CompleteIxAy = AyAddAp(PartialIxAy, AyMinus(I, PartialIxAy))
End Property

Property Get Pop(Ay)
Pop = AyLasEle(Ay)
AyRmvLasNEle Ay
End Property

Property Get Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Property

Property Get UB&(Ay)
UB = Sz(Ay) - 1
End Property

Property Get UIntSeq(U&, Optional IsFmOne As Boolean) As Integer()
Dim O%(): ReDim O(U)
Dim J&
If IsFmOne Then
    For J = 0 To U
        O(J) = J + 1
    Next
Else
    For J = 0 To U
        O(J) = J
    Next
End If
UIntSeq = O
End Property

Property Get ULngSeq(U&, Optional IsFmOne As Boolean) As Long()
Dim O&()
ReDim O(U)
Dim J&
If IsFmOne Then
    For J = 0 To U
        O(J) = J + 1
    Next
Else
    For J = 0 To U
        O(J) = J
    Next
End If
ULngSeq = O
End Property

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

Sub AyIxAy_Asg(Ay, IxAy&(), ParamArray OAp())
Dim J%
For J = 0 To UB(IxAy)
    OAp(J) = Ay(IxAy(J))
Next
End Sub

Sub AyRmvLasNEle(Ay, Optional NEle% = 1)
ReDim Preserve Ay(UB(Ay) - NEle)
End Sub

Sub AyWrt(Ay, Ft)
StrWrt JnCrLf(Ay), Ft
End Sub

Sub Push(OAy, M)
Dim N&: N = Sz(OAy)
ReDim Preserve OAy(N)
If IsObject(M) Then
    Set OAy(N) = M
Else
    OAy(N) = M
End If
End Sub

Sub PushAp(O, ParamArray Ap())
Dim Av(), I: Av = Ap
For Each I In Av
    Push O, I
Next
End Sub

Sub PushAy(OAy, Ay)
If AyIsEmp(Ay) Then Exit Sub
Dim I
For Each I In Ay
    Push OAy, I
Next
End Sub

Sub PushNoDup(O, M)
If Not AyHas(O, M) Then Push O, M
End Sub

Sub PushNoDupAy(O, Ay)
Dim I
If AyIsEmp(Ay) Then Exit Sub
For Each I In Ay
    PushNoDup O, I
Next
End Sub

Sub PushNonEmp(O, M)
If IsEmp(M) Then Exit Sub
Push O, M
End Sub

Sub PushObj(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
Set O(N) = P
End Sub

Sub PushObjAy(O, Ay)
Dim J&
For J = 0 To UB(Ay)
    PushObj O, Ay(J)
Next
End Sub

Sub PushOy(O, Oy)
If AyIsEmp(Oy) Then Exit Sub
Dim M
For Each M In Oy
    PushObj O, M
Next
End Sub

Sub ReSz(Ay, U&)
If U < 0 Then
    Erase Ay
Else
    ReDim Preserve Ay(U)
End If
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

Private Property Get AySrt__Ix&(Ay, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ay
        If V > I Then AySrt__Ix = O: Exit Property
        O = O + 1
    Next
    AySrt__Ix = O
    Exit Property
End If
For Each I In Ay
    If V < I Then AySrt__Ix = O: Exit Property
    O = O + 1
Next
AySrt__Ix = O
End Property

Sub ZZ__Tst()
ZZ_AyTrim
End Sub

Private Property Get AySrtInToIxAy_Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then AySrtInToIxAy_Ix& = O: Exit Property
        O = O + 1
    Next
    AySrtInToIxAy_Ix& = O
    Exit Property
End If
For Each I In Ix
    If V < A(I) Then AySrtInToIxAy_Ix& = O: Exit Property
    O = O + 1
Next
AySrtInToIxAy_Ix& = O
End Property

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

Private Sub ZZ_AyFmTo_Brk()
Dim A(): A = Array(1, 2, 3, 4)
Dim M As FmTo: M = FmTo(1, 2)
Dim Act(): Act = AyFmTo_Brk(A, M)
Ass Sz(Act) = 2
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

Private Sub ZZ_AyMap_Sy()
Dim Ay$(): Ay = AyMap_Sy(Array("skldfjdf", "aa"), "RmvFstChr")
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
FmTo.FmIx = 1
FmTo.ToIx = 2
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

Private Sub ZZ_IxAy_IsParitial_of_0toU()
Ass IxAy_IsParitial_of_0toU(ApLngAy(1, 2, 3, 4), 4) = True
Ass IxAy_IsParitial_of_0toU(ApLngAy(0, 1, 2, 3, 4), 4) = True
Ass IxAy_IsParitial_of_0toU(ApLngAy(1, 1, 3, 4), 4) = False
Ass IxAy_IsParitial_of_0toU(ApLngAy(5, 3, 4), 4) = False
End Sub
