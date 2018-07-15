Attribute VB_Name = "VbAy"
Option Explicit

Property Get AyRmvEle(A, Ele)
Dim Ix&: Ix = AyIx(A, Ele): If Ix = -1 Then AyRmvEle = A: Exit Property
AyRmvEle = AyRmvEleAt(A, AyIx(A, Ele))
End Property

Property Get AyRmvEleAt(A, Optional At&)
AyRmvEleAt = AyWhExclAtCnt(A, At)
End Property

Property Get AyWhExclAtCnt(A, At&, Optional Cnt& = 1)
If Cnt <= 0 Then AyWhExclAtCnt = A: Exit Property
Dim U&: U = UB(A)
If At > U Then Stop
If At < 0 Then Stop
If U = 0 Then AyWhExclAtCnt = A: Exit Property
Dim O: O = A
Dim J&
For J = At To U - Cnt
    O(J) = O(J + Cnt)
Next
ReDim Preserve O(U - Cnt)
AyWhExclAtCnt = O
End Property

Function ApIntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
ApIntAy = AyIntAy(Av)
End Function

Function ApLngAy(ParamArray Ap()) As Long()
Dim Av(): Av = Ap
ApLngAy = AyLngAy(Av)
End Function

Function ApSngAy(ParamArray Ap()) As Single()
Dim Av(): Av = Ap
ApSngAy = AySngAy(Av)
End Function

Function ApSy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
ApSy = AySy(Av)
End Function

Sub AssChk(Chk$())
If AyIsEmp(Chk) Then Exit Sub
AyBrw Chk
Stop
End Sub

Function AyAdd(Ay1, Ay2)
Dim O: O = Ay1
PushAy O, Ay2
AyAdd = O
End Function

Function AyAddAp(A, ParamArray AyAp())
Dim Av(): Av = AyAp
Dim Ay
Dim O: O = A
For Each Ay In Av
    PushAy O, Ay
Next
AyAddAp = O
End Function

Function AyAddPfx(A, Pfx) As String()
Dim O$(), U&, J&
U = UB(A)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & A(J)
Next
AyAddPfx = O
End Function

Function AyAddPfxSfx(A, Pfx, Sfx) As String()
Dim O$(), U&, J&
U = UB(A)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = Pfx & A(J) & Sfx
Next
AyAddPfxSfx = O
End Function

Function AyAddSfx(A, Sfx) As String()
Dim O$(), U&, J&
U = UB(A)
If U = -1 Then Exit Function
ReDim Preserve O(U)
For J = 0 To U
    O(J) = A(J) & Sfx
Next
AyAddSfx = O
End Function

Function AyAlignL(A) As String()
If AyIsEmp(A) Then Exit Function
Dim W%: W = AyWdt(A)
Dim O$(), I
For Each I In A
    Push O, AlignL(CStr(I), W)
Next
AyAlignL = O
End Function

Sub AyAsg(A, ParamArray OAp())
Dim V
Dim J%
For Each V In A
    If Not IsMissing(OAp(J)) Then
        OAp(J) = A(J)
    End If
    J = J + 1
Next
End Sub

Function AyAsgAy(A, OIntoAy)
If TypeName(A) = TypeName(OIntoAy) Then
    OIntoAy = A
    AyAsgAy = OIntoAy
    Exit Function
End If
If AyIsEmp(A) Then
    Erase OIntoAy
    AyAsgAy = OIntoAy
    Exit Function
End If
Dim U&
    U = UB(A)
ReDim OIntoAy(U)
Dim I, J&
For Each I In A
    Asg I, OIntoAy(J)
    J = J + 1
Next
AyAsgAy = OIntoAy
End Function

Function AyBoolAy(A) As Boolean()
AyBoolAy = AyAsgAy(A, Emp.BoolAy)
End Function

Function AyBrk3ByIx(A, FmIx&, ToIx&)
AyBrk3ByIx = AyFmTo_Brk(A, NewFmTo(FmIx, ToIx))
End Function

Function AyBrw(A, Optional Fnn$)
Dim T$
T = TmpFt("AyBrw", Fnn)
AyWrt A, T
FtBrw T
End Function

Function AyBytAy(A) As Byte()
AyBytAy = AyAsgAy(A, Emp.BytAy)
End Function

Function AyC1Dry(A, C) As Variant()
'C1Dry is Dry with 2 column and Col1 is const
Dim U&, J&
U = UB(A)
Dim O()
ReSz O, U
For J = 0 To U
    O(J) = Array(C, A(J))
Next
AyC1Dry = O
End Function

Function AyC2Dry(A, C) As Variant()
'C2Dry is Dry with 2 column and Col1 is const
Dim U&, J&
U = UB(A)
Dim O()
ReSz O, U
For J = 0 To U
    O(J) = Array(A(J), C)
Next
AyC2Dry = O
End Function

Function AyDblAy(A) As Double()
AyDblAy = AyAsgAy(A, Emp.DblAy)
End Function

Function AyDic(A, Optional V = True) As Dictionary
Dim O As New Dictionary, I
If Not AyIsEmp(A) Then
    For Each I In A
        O.Add I, V
    Next
End If
Set AyDic = O
End Function

Sub AyDmp(A, Optional WithIx As Boolean)
If AyIsEmp(A) Then Exit Sub
Dim I
If WithIx Then
    Dim J&
    For Each I In A
        Debug.Print J; ": "; I
        J = J + 1
    Next
Else
    For Each I In A
        Debug.Print I
    Next
End If

End Sub

Function AyDrs(A) As Drs
AyDrs = NewDrs("Itm", AyDry(A))
End Function

Function AyDry(A) As Variant()
Dim O(), J&
Dim U&: U = UB(A)
ReSz O, U
For J = 0 To U
    O(J) = Array(A(J))
Next
AyDry = O
End Function

Function AyDteAy(A) As Date()
AyDteAy = AyAsgAy(A, Emp.DteAy)
End Function

Function AyDupAy(A)
'Return Array of element of {Ay} for which has 2 or more value in {Ay}
Dim OAy: OAy = A: Erase OAy
If Not AyIsEmp(A) Then
    Dim Uniq: Uniq = OAy
    Dim V
    
    For Each V In A
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
If VarIsEmp(O1) Then Exit Function
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

Function AyFmTo_Brk(A, B As FmTo)
Ass FmTo_HasU(B, UB(A))
Dim O(2)
O(0) = AyWhFmTo(A, NewFmTo(0, B.FmIx - 1))
O(1) = AyWhFmTo(A, B)
O(2) = AyWhFmTo(A, NewFmTo(B.FmIx + 1, UB(A)))
AyFmTo_Brk = O
End Function

Function AyFstNEle(A, N&)
Dim O: O = A
ReDim Preserve O(N - 1)
AyFstNEle = O
End Function

Function AyGpDry(A) As Variant()
If AyIsEmp(A) Then Exit Function
Dim O(), I
For Each I In A
    AyGpDry__Upd O, I
Next
AyGpDry = O
End Function

Function AyHas(A, Itm) As Boolean
If AyIsEmp(A) Then Exit Function
Dim I
For Each I In A
    If I = Itm Then AyHas = True: Exit Function
Next
End Function

Function AyHasDupEle(A) As Boolean
If AyIsEmp(A) Then Exit Function
Dim Pool: Pool = A: Erase Pool
Dim I
For Each I In A
    If AyHas(Pool, I) Then AyHasDupEle = True: Exit Function
    Push Pool, I
Next
End Function

Function AyHasNegOne(A) As Boolean
Dim V
If AyIsEmp(A) Then Exit Function
For Each V In A
    If V = -1 Then AyHasNegOne = True: Exit Function
Next
End Function

Function AyHasSubAy(A, SubAy) As Boolean
Const CSub$ = "AyHasSubAy"
If AyIsEmp(A) Then Exit Function
If AyIsEmp(SubAy) Then Er CSub, "{SubAy} is empty", SubAy
Dim I
For Each I In SubAy
    If Not AyHas(A, I) Then Exit Function
Next
End Function

Function AyIncNForEachEle(A, Optional N& = 1)
Dim O: O = A
Dim J&
For J = 0 To UB(A)
    O(J) = O(J) + N
Next
AyIncNForEachEle = O
End Function

Function AyIns(A, Optional Ele, Optional At&)
Ass IsArray(A)
Const CSub$ = "AyIns"
Dim N&: N = Sz(A)
If 0 > At Or At > N Then Er CSub, "{At} is outside {Ay-UB}", At, UB(A)
Dim O
    O = A
    ReDim Preserve O(N)
    Dim J&
    For J = N To At + 1 Step -1
        Asg O(J - 1), O(J)
    Next
    O(At) = Ele
AyIns = O
End Function

Function AyIntAy(A) As Integer()
AyIntAy = AyAsgAy(A, Emp.IntAy)
End Function

Function AyIntersect(Ay1, Ay2)
Dim O: O = Ay1: Erase O
If AyIsEmp(Ay1) Then GoTo X
If AyIsEmp(Ay2) Then GoTo X
Dim V
For Each V In Ay1
    If AyHas(Ay2, V) Then Push O, V
Next
X:
AyIntersect = O
End Function

Function AyIsAllEleHasVal(A) As Boolean
If AyIsEmp(A) Then Exit Function
Dim I
For Each I In A
    If VarIsEmp(I) Then Exit Function
Next
AyIsAllEleHasVal = True
End Function

Function AyIsAllEq(A) As Boolean
If AyIsEmp(A) Then AyIsAllEq = True: Exit Function
Dim T: T = A(0)
Dim J&
For J = 1 To UB(A)
    If A(J) = T Then Exit Function
Next
AyIsAllEq = True
End Function

Function AyIsEmp(V) As Boolean
AyIsEmp = (Sz(V) = 0)
End Function

Function AyIsEq(A1, A2) As Boolean
Dim U&: U = UB(A1): If U <> UB(A2) Then Exit Function
Dim J&
For J = 0 To U
   If A1(J) <> A2(J) Then Exit Function
Next
AyIsEq = True
End Function

Function AyIsSamSz(A1, A2) As Boolean
AyIsSamSz = Sz(A1) = Sz(A2)
End Function

Function AyIsSrt(Ay) As Boolean
Dim J&
For J = 0 To UB(Ay) - 1
   If Ay(J) > Ay(J + 1) Then Exit Function
Next
AyIsSrt = True
End Function

Function AyIx&(A, Itm)
Dim J&
For J = 0 To UB(A)
    If A(J) = Itm Then AyIx = J: Exit Function
Next
AyIx = -1
End Function

Function AyIxAy(A, SubAy, Optional ChkNotFound As Boolean, Optional SkipNotFound As Boolean) As Long()
If AyIsEmp(SubAy) Then Exit Function
Dim O&()
Dim U&: U = UB(SubAy)
Dim J&, Ix&
If SkipNotFound Then
    For J = 0 To U
        Ix = AyIx(A, SubAy(J))
        If Ix >= 0 Then
            Push O, Ix
        End If
    Next
Else
    ReDim O(U)
    For J = 0 To U
        O(J) = AyIx(A, SubAy(J))
    Next
End If
If Not SkipNotFound And ChkNotFound Then
    AyIxAy__ChkNotFound O, A, SubAy
End If
AyIxAy = O
End Function

Sub AyIxAy_Asg(A, IxAy&(), ParamArray OAp())
Dim J%
For J = 0 To UB(IxAy)
    OAp(J) = A(IxAy(J))
Next
End Sub

Function AyLasEle(A)
AyLasEle = A(UB(A))
End Function

Function AyLngAy(A) As Long()
AyLngAy = AyAsgAy(A, Emp.LngAy)
End Function

Function AyMap(A, MthNm$, ParamArray Ap()) As Variant()
If AyIsEmp(A) Then Exit Function
Dim Av(): Av = Ap
Av = AyIns(Av)
Dim I, J&
Dim O()
Dim U&: U = UB(A)
    ReDim O(U)
For Each I In A
    Asg I, Av(0)
    Asg RunAv(MthNm, Av), O(J)
    J = J + 1
Next
AyMap = O
End Function

Function AyMapAsgAy(A, OAy, MthNm$, ParamArray Ap())
If AyIsEmp(A) Then Exit Function
Dim Av(): Av = Ap
Av = AyIns(Av)
Dim I, J&
Dim O
O = OAy
Erase O
Dim U&: U = UB(A)
    ReDim O(U)
For Each I In A
    Asg I, Av(0)
    Asg RunAv(MthNm, Av), O(J)
    J = J + 1
Next
AyMapAsgAy = O
End Function

Function AyMapAsgSy(A, MthNm$, ParamArray Ap()) As String()
If AyIsEmp(A) Then Exit Function
Dim Av(): Av = Ap
If AyIsEmp(Av) Then
    AyMapAsgSy = AyMap_Sy(A, MthNm)
    Exit Function
End If
Dim I, J&
Dim O$()
    ReDim O(UB(A))
    Av = AyIns(Av)
    For Each I In A
        Asg I, Av(0)
        Asg RunAv(MthNm, Av), O(J)
        J = J + 1
    Next
AyMapAsgSy = O
End Function

Function AyMapInto(Ay, Obj, GetNm$, OIntoAy)
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
End Function

Function AyMap_Lng(A, MapMthNm$) As Long()
AyMap_Lng = AyLngAy(AyMap(A, MapMthNm))
End Function

Function AyMap_Sy(A, MapMthNm$) As String()
AyMap_Sy = AySy(AyMap(A, MapMthNm))
End Function

Function AyMax(A)
If AyIsEmp(A) Then Exit Function
Dim O, I
For Each I In A
    If I > O Then O = I
Next
AyMax = O
End Function

Sub AyMdyByRmvFstEle(OAy)
Dim J&
For J = 0 To UB(OAy) - 1
    OAy(J) = OAy(J + 1)
Next
AyMdyByRmvLasEle OAy
End Sub

Sub AyMdyByRmvLasEle(OAy)
ReDim Preserve OAy(UB(OAy) - 1)
End Sub

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
Dim O: O = A
Dim Av(): Av = AyAp
Dim Ay1, V
For Each Ay1 In Av
    If AyIsEmp(O) Then GoTo X
    O = AyMinus(O, Ay1)
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

Sub AyPair_EqChk(Ay1, Ay2, Optional Ay1Nm$ = "Exp", Optional Ay2Nm$ = "Act")
AssChk AyEqChk(Ay1, Ay2, Ay1Nm, Ay2Nm)
End Sub

Function AyQuote(A, QuoteStr$) As String()
If AyIsEmp(A) Then Exit Function
Dim U&: U = UB(A)
Dim O$()
    ReDim O(U)
    Dim J&
    Dim Q1$, Q2$
    With BrkQuote(QuoteStr)
        Q1 = .S1
        Q2 = .S2
    End With
    For J = 0 To U
        O(J) = Q1 & A(J) & Q2
    Next
AyQuote = O
End Function

Function AyQuoteDbl(A) As String()
AyQuoteDbl = AyQuote(A, """")
End Function

Function AyQuoteSng(A) As String()
AyQuoteSng = AyQuote(A, "'")
End Function

Function AyQuoteSqBkt(A) As String()
AyQuoteSqBkt = AyQuote(A, "[]")
End Function

Function AyRTrim(A) As String()
If AyIsEmp(A) Then Exit Function
Dim O$(), I
For Each I In A
    Push O, RTrim(I)
Next
AyRTrim = O
End Function

Function AyReOrd(A, PartialIxAy&())
Dim I&()
    I = PartialIxAy_CompleteIxAy(PartialIxAy, UB(A))
Dim O
    O = A: Erase O
    Dim J&
    For J = 0 To UB(I)
        Push O, A(I(J))
    Next
AyReOrd = O
End Function

Function AyRmvEmp(A)
If AyIsEmp(A) Then AyRmvEmp = A: Exit Function
Dim O: O = A: Erase O
Dim I
For Each I In A
    If Not VarIsEmp(I) Then Push O, I
Next
AyRmvEmp = O
End Function

Function AyRmvEmpEleAtEnd(A)
Dim LasU&, U&
Dim O: O = A
For LasU = UB(A) To 0 Step -1
    If Not VarIsEmp(O(LasU)) Then
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

Function AyRmvFmTo(A, FmTo As FmTo)
Dim O
    O = A
    If Not (IsEmpFmTo(FmTo) Or AyIsEmp(A)) Then
        Dim FmI&, ToI&
        FmI = FmTo.FmIx
        ToI = FmTo.ToIx
        Dim I&, J&, U&
        U = UB(A)
        I = 0
        For J = ToI + 1 To U
            O(FmI + I) = O(J)
            I = I + 1
        Next
        ReDim Preserve O(U - FmTo_N(FmTo))
    End If
AyRmvFmTo = O
End Function

Function AyRmvFstEle(A)
AyRmvFstEle = AyRmvEleAt(A)
End Function

Function AyRmvLasEle(A)
AyRmvLasEle = AyRmvEleAt(A, UB(A))
End Function

Sub AyRmvLasNEle(A, Optional NEle% = 1)
ReDim Preserve A(UB(A) - NEle)
End Sub

Function AyRmvPfx(A, Pfx) As String()
If AyIsEmp(A) Then Exit Function
Dim U&: U = UB(A)
Dim O$()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = RmvPfx(A(J), Pfx)
Next
AyRmvPfx = O
End Function

Function AyRpl(A, FmTo As FmTo, AySeg)
Dim Ay()
    Ay = AyFmTo_Brk(A, FmTo)
Dim O
    O = A(0): Erase O
    PushAy O, AySeg
    PushAy O, A(2)
AyRpl = O
End Function

Function AyShift(OAy)
AyShift = OAy(0)
OAy = AyRmvFstEle(OAy)
End Function

Function AySngAy(A) As Single()
AySngAy = AyAsgAy(A, Emp.SngAy)
End Function

Function AySqH(A) As Variant()
If AyIsEmp(A) Then Exit Function
Dim O(), C%
ReDim O(1 To 1, 1 To Sz(A))
C = 0
Dim V
For Each V In A
    C = C + 1
    O(1, C) = V
Next
AySqH = O
End Function

Function AySqV(A) As Variant()
If AyIsEmp(A) Then Exit Function
Dim O(), R&
ReDim O(1 To Sz(A), 1 To 1)
R = 0
Dim V
For Each V In A
    R = R + 1
    O(R, 1) = V
Next
AySqV = O
End Function

Function AySrt(A, Optional Des As Boolean)
If AyIsEmp(A) Then AySrt = A: Exit Function
Dim Ix&, V, J&
Dim O: O = A: Erase O
Push O, A(0)
For J = 1 To UB(A)
    O = AyIns(O, A(J), AySrt__Ix(O, A(J), Des))
Next
AySrt = O
End Function

Function AySrtInToIxAy(A, Optional Des As Boolean) As Long()
If AyIsEmp(A) Then Exit Function
Dim Ix&, V, J&
Dim O&():
Push O, 0
For J = 1 To UB(A)
    O = AyIns(O, J, AySrtInToIxAy_Ix(O, A, A(J), Des))
Next
AySrtInToIxAy = O
End Function

Function AySy(A) As String()
AySy = AyAsgAy(A, Emp.Sy)
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

Function AyUniq(A)
Dim O: O = A: Erase O
Dim V
If Not AyIsEmp(A) Then
    For Each V In A
        PushNoDup O, V
    Next
End If
AyUniq = O
End Function

Function AyVSq(A)
Dim O
Dim N&
N = Sz(A)
ReDim O(1 To N, 1 To 1)
Dim J&
For J = 1 To N
    O(J, 1) = A(J - 1)
Next
AyVSq = O
End Function

Function AyWdt%(A)
If AyIsEmp(A) Then Exit Function
Dim O%, I
For Each I In A
    O = Max(O, Len(I))
Next
AyWdt = O
End Function

Function AyWh(A, FmIx&, ToIx&)
Dim O: O = A: Erase O
AyWh = O
If AyIsEmp(A) Then Exit Function
If FmIx < 0 Then Exit Function
If ToIx < 0 Then Exit Function
Dim J&
For J = FmIx To ToIx
    Push O, A(J)
Next
AyWh = O
End Function

Function AyWhDist(A)
Dim O: O = A: Erase O
Dim I
For Each I In A
    PushNoDup O, I
Next
AyWhDist = O
End Function

Function AyWhDup(A)
Dim O: O = A: Erase O
Dim GpDry(): GpDry = AyGpDry(A)
If AyIsEmp(GpDry) Then AyWhDup = O: Exit Function
Dim Dr
For Each Dr In GpDry
    If Dr(1) > 1 Then Push O, Dr(0)
Next
AyWhDup = O
End Function

Function AyWhExclIxAy(A, IxAy)
'IxAy holds index if A to be remove.  It has been sorted else will be stop
Ass AyIsSrt(A)
Ass AyIsSrt(IxAy)
Dim J&
Dim O: O = A
For J = UB(IxAy) To 0 Step -1
    O = AyRmvEleAt(O, CLng(IxAy(J)))
Next
AyWhExclIxAy = O
End Function

Function AyWhFm(A, FmIx&)
Dim O: O = A: Erase O
If 0 <= FmIx And FmIx <= UB(A) Then
    Dim J&
    For J = FmIx To UB(A)
        Push O, A(J)
    Next
End If
AyWhFm = O
End Function

Function AyWhFmTo(A, FmTo As FmTo)
AyWhFmTo = AyWh(A, FmTo.FmIx, FmTo.ToIx)
End Function

Function AyWhIxAy(A, IxAy, Optional CrtEmpColIfReqFldNotFound As Boolean)
'Return a subset of {Ay} by {IxAy}
Ass IsArray(A)
Ass IsArray(IxAy)
Dim O
    O = A: Erase O
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
            If Not CrtEmpColIfReqFldNotFound Then
                Er "AyWhIxAy", "Given {IxAy} contains -1", IxAy
            End If
        Else
            If IsObject(A(Ix)) Then
                Set O(J) = A(Ix)
            Else
                O(J) = A(Ix)
            End If
        End If
    Next
X:
AyWhIxAy = O
End Function

Function AyWhLik(A, Lik$) As String()
If AyIsEmp(A) Then Exit Function
Dim O$()
Dim I
For Each I In A
    If I Like Lik Then Push O, I
Next
AyWhLik = O
End Function

Function AyWhLikAy(A, LikAy$()) As String()
If AyIsEmp(A) Then Exit Function
If AyIsEmp(LikAy) Then Exit Function
Dim I, Lik, O$()
For Each I In A
    For Each Lik In LikAy
        If I Like Lik Then
            Push O, I
            Exit For
        End If
    Next
Next
AyWhLikAy = O
End Function

Function AyWhMulEle(A)
'Return Set of Element as array in {Ay} having 2 or more element
Dim Dry(): Dry = AyGpDry(A)
Dim O: O = A: Erase O
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

Function AyWhPatnIx(A, Patn$) As Long()
If AyIsEmp(A) Then Exit Function
Dim I, O&(), J&
Dim R As Re
Set R = Re(Patn)
For Each I In A
    If R.Tst(I) Then Push O, J
    J = J + 1
Next
AyWhPatnIx = O
End Function

Function AyWhSfx(A, Sfx$) As String()
If AyIsEmp(A) Then Exit Function
Dim O$()
Dim I
For Each I In A
    If HasSfx(CStr(I), Sfx) Then Push O, I
Next
AyWhSfx = O
End Function

Function AyWhSngEle(A)
'Return Set of Element as array in {Ay} having 2 or more element
Dim Dry(): Dry = AyGpDry(A)
Dim O: O = A: Erase O
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

Function AyWh_ByIxAy(A, IxAy)
Dim O: O = A: Erase O
Dim J%
For J = 0 To UB(IxAy)
    Push O, A(IxAy(J))
Next
AyWh_ByIxAy = O
End Function

Function AyWh_ByMth(A, WhMthNm$, ParamArray Ap())
Dim O: O = A: Erase O
Dim I
Dim Av()
    Av = Ap
    Av = AyIns(Av)
For Each I In A
    Asg I, Av(0)
    If RunAv(WhMthNm, Av) Then
        Push O, I
    End If
Next
AyWh_ByMth = O
End Function

Function AyWh_ByPatn(A, Patn$) As String()
If AyIsEmp(A) Then Exit Function
Dim I, O$()
Dim R As Re
Set R = Re(Patn)
For Each I In A
    If R.Tst(I) Then Push O, I
Next
AyWh_ByPatn = O
End Function

Function AyWh_ByPfx(A, Pfx$) As String()
If AyIsEmp(A) Then Exit Function
Dim O$()
Dim I
For Each I In A
    If HasPfx(CStr(I), Pfx) Then Push O, I
Next
AyWh_ByPfx = O
End Function

Sub AyWrt(A, Ft)
StrWrt JnCrLf(A), Ft
End Sub

Function AyWs(A, Optional WsNm$, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs(WsNm, Vis)
SqRg AyVSq(A), WsA1(O)
Set AyWs = O
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

Function CvAy(V) As Variant()
CvAy = V
End Function

Function CvSy(V) As String()
CvSy = V
End Function

Function FnyIxAy(A$(), SubFny0) As Integer()
Dim SubFny$(): SubFny = DftNy(SubFny0)
If AyIsEmp(SubFny) Then Stop
Dim O%(), U&, J%
U = UB(SubFny)
ReSz O, U
For J = 0 To U
    O(J) = AyIx(A, SubFny(J))
    If O(J) = -1 Then Stop
Next
End Function

Function IntAy_ByU(U&) As Integer()
If 0 > U Then Exit Function
Dim O%()
ReDim O(U)
Dim J&
For J = 0 To U
    O(J) = J
Next
IntAy_ByU = O
End Function

Function IxAy_IsAllGE0(IxAy&()) As Boolean
Dim J&
For J = 0 To UB(IxAy)
    If IxAy(J) = -1 Then Exit Function
Next
IxAy_IsAllGE0 = True
End Function

Function IxAy_IsParitial_of_0toU(IxAy, U&) As Boolean
Const CSub$ = "Ass IxAy_IsParitial_of_0toU"
Const Msg$ = "{IxAy} is not PartialIx-of-{U}." & _
"|PartialIxAy-Of-U is defined as:" & _
"|It should be Lng()" & _
"|It should have 0 to U elements" & _
"|It should have each element of value between 0 and U" & _
"|It should have no dup element" & _
"|All elements should have value equal or less than U"

If Not VarIsLngAy(IxAy) Then Exit Function
If AyIsEmp(IxAy) Then IxAy_IsParitial_of_0toU = True: Exit Function
If AyHasDupEle(IxAy) Then Exit Function
Dim I
For Each I In IxAy
   If 0 > I Or I > U Then Exit Function
Next
IxAy_IsParitial_of_0toU = True
End Function

Function LinesEndTrim$(A$)
LinesEndTrim = JnCrLf(LyEndTrim(SplitCrLf(A)))
End Function

Function LyEndTrim(A$()) As String()
If AyIsEmp(A) Then Exit Function
If Not Lin(AyLasEle(A)).IsEmp Then LyEndTrim = A: Exit Function
Dim J%
For J = UB(A) To 0 Step -1
    If Not Lin(A(J)).IsEmp Then
        Dim O$()
        O = A
        ReDim Preserve O(J)
        LyEndTrim = O
        Exit Function
    End If
Next
End Function

Function LyRmv2Dash(Ly$()) As String()
If AyIsEmp(Ly) Then Exit Function
Dim O$(), I
For Each I In Ly
    Push O, Rmv2Dash(CStr(I))
Next
LyRmv2Dash = O
End Function

Function NewDblAy(ParamArray Ap()) As Double()
Dim Av(): Av = Ap
NewDblAy = AyDblAy(Av)
End Function

Function NewDteAy(ParamArray Ap()) As Date()
Dim Av(): Av = Ap
NewDteAy = AyDteAy(Av)
End Function

Function NewIntSeq(N&, Optional IsFmOne As Boolean) As Integer()
Dim O%(): ReDim O(N - 1)
Dim J&
If IsFmOne Then
    For J = 0 To N - 1
        O(J) = J + 1
    Next
Else
    For J = 0 To N - 1
        O(J) = J
    Next
End If
NewIntSeq = O
End Function

Function NewIxAy(U&) As Long()
Dim O&()
    ReDim O(U)
    Dim J&
    For J = 0 To U
        O(J) = J
    Next
NewIxAy = O
End Function

Function NewSy(U&) As String()
Dim O$()
If U > 0 Then ReDim O(U)
NewSy = O
End Function

Function PartialIxAy_CompleteIxAy(PartialIxAy&(), U&) As Long()
'Des:Make a complete-IxAy-of-U by partialIxAy
'Des:A complete-IxAy-Of-U is defined as
'Des:it has (U+1)-elements,
'Des:it does not have dup element
'Des:it has all element of value between 0 and U
Ass IxAy_IsParitial_of_0toU(PartialIxAy, U)
Dim I&(): I = NewIxAy(U)
PartialIxAy_CompleteIxAy = AyAddAp(PartialIxAy, AyMinus(I, PartialIxAy))
End Function

Function Pop(A)
Pop = AyLasEle(A)
AyRmvLasNEle A
End Function

Sub Push(O, M)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub

Sub PushAp(O, ParamArray Ap())
Dim Av(), I: Av = Ap
For Each I In Av
    Push O, I
Next
End Sub

Sub PushAy(O, A)
If AyIsEmp(A) Then Exit Sub
Dim I
For Each I In A
    Push O, I
Next
End Sub

Sub PushItmAy(O, Itm, Ay)
Push O, Itm
PushAy O, Ay
End Sub

Sub PushNoDup(O, M)
If Not AyHas(O, M) Then Push O, M
End Sub

Sub PushNoDupAy(O, A)
Dim I
If AyIsEmp(A) Then Exit Sub
For Each I In A
    PushNoDup O, I
Next
End Sub

Sub PushNonEmp(O, M)
If VarIsEmp(M) Then Exit Sub
Push O, M
End Sub

Sub PushObj(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
Set O(N) = P
End Sub

Sub PushOy(O, Oy)
If AyIsEmp(Oy) Then Exit Sub
Dim M
For Each M In Oy
    PushObj O, M
Next
End Sub

Sub ReSz(A, U&)
If U < 0 Then
    Erase A
Else
    ReDim Preserve A(U)
End If
End Sub

Function SyAddAp(ParamArray Ap()) As String()
Dim Av(): Av = Ap
Dim O$(), I
For Each I In Av
    If IsStr(I) Then
        Push O, I
    Else
        PushAy O, I
    End If
Next
End Function

Function SyIsAllEleHasPfx(A$(), Pfx$) As Boolean
If AyIsEmp(A) Then Exit Function
Dim I
For Each I In A
   If Not HasPfx(CStr(I), Pfx) Then Exit Function
Next
SyIsAllEleHasPfx = True
End Function

Sub SyMdyByRmvFstChr(OSy$())
Dim J&
For J = 0 To UB(OSy)
    OSy(J) = RmvFstChr(OSy(J))
Next
End Sub

Function SyRmvLasChr(A$()) As String()
SyRmvLasChr = AyMap_Sy(A, "RmvLasChr")
End Function

Function SyTrim(Sy$()) As String()
If AyIsEmp(Sy) Then Exit Function
Dim U&
    U = UB(Sy)
Dim O$()
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = Trim(Sy(J))
    Next
SyTrim = O
End Function

Function Sz&(A)
On Error Resume Next
Sz = UBound(A) + 1
End Function

Function UB&(A)
UB = Sz(A) - 1
End Function

Sub AyAdd__Tst()
Dim Act(), Exp(), Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyAdd(Ay1, Ay2)
Exp = Array(1, 2, 2, 2, 4, 5, 2, 2)
AyPair_EqChk Exp, Act
AyPair_EqChk Ay1, Array(1, 2, 2, 2, 4, 5)
AyPair_EqChk Ay2, Array(2, 2)
End Sub

Private Sub AyEqChk__Tst()
AyDmp AyEqChk(Array(1, 2, 3, 3, 4), Array(1, 2, 3, 4, 4))
End Sub

Sub AyFmTo_Brk__Tst()
Dim A(): A = Array(1, 2, 3, 4)
Dim M As FmTo: M = NewFmTo(1, 2)
Dim Act(): Act = AyFmTo_Brk(A, M)
Ass Sz(Act) = 2
Ass AyIsEq(Act(0), Array(1))
Ass AyIsEq(Act(1), Array(2, 3))
Ass AyIsEq(Act(2), Array(4))
End Sub

Private Sub AyGpDry__Tst1()
Dim A$(): A = SplitSpc("a a a b c b")
Dim Act(): Act = AyGpDry(A)
Dim Exp(): Exp = Array(Array("a", 3), Array("b", 2), Array("c", 1))
AssEqDry Act, Exp
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

Sub AyHasDupEle__Tst()
Ass AyHasDupEle(Array(1, 2, 3, 4)) = False
Ass AyHasDupEle(Array(1, 2, 3, 4, 4)) = True
End Sub

Sub AyIntAy__Tst()
Dim Act%(): Act = AyIntAy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
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

Sub AyMap_Sy__Tst()
Dim Ay$(): Ay = AyMap_Sy(Array("skldfjdf", "aa"), "RmvFstChr")
Stop
End Sub

Private Sub AyMap__Tst()
Dim Act: Act = AyMap(Array(1, 2, 3, 4), "Mul2")
Ass Sz(Act) = 4
Ass Act(0) = 2
Ass Act(1) = 4
Ass Act(2) = 6
Ass Act(3) = 8
End Sub

Sub AyMinus__Tst()
Dim Act(), Exp()
Dim Ay1(), Ay2()
Ay1 = Array(1, 2, 2, 2, 4, 5)
Ay2 = Array(2, 2)
Act = AyMinus(Ay1, Ay2)
Exp = Array(1, 2, 4, 5)
AyPair_EqChk Exp, Act
'
Act = AyMinusAp(Array(1, 2, 2, 2, 4, 5), Array(2, 2), Array(5))
Exp = Array(1, 2, 4)
AyPair_EqChk Exp, Act
End Sub

Private Sub AyRmvEmpEleAtEnd__Tst()
Dim A: A = Array(Empty, Empty, Empty, 1, Empty, Empty)
Dim Act: Act = AyRmvEmpEleAtEnd(A)
Ass Sz(Act) = 4
Ass Act(3) = 1
End Sub

Private Sub AyRmvFmTo__Tst()
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

Private Sub AySrtInToIxAy__Tst()
Dim A: A = Array("A", "B", "C", "D", "E")
AyPair_EqChk Array(0, 1, 2, 3, 4), AySrtInToIxAy(A)
AyPair_EqChk Array(4, 3, 2, 1, 0), AySrtInToIxAy(A, True)
End Sub

Private Function AySrt__Ix&(A, V, Des As Boolean)
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

Private Sub AySrt__Tst()
Dim Exp, Act
Dim A
A = Array(1, 2, 3, 4, 5): Exp = A:                   Act = AySrt(A):       AyPair_EqChk Exp, Act
A = Array(1, 2, 3, 4, 5): Exp = Array(5, 4, 3, 2, 1): Act = AySrt(A, True): AyPair_EqChk Exp, Act
A = Array(":", "~", "P"): Exp = Array(":", "P", "~"): Act = AySrt(A):       AyPair_EqChk Exp, Act
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
AyPair_EqChk Exp, Act
End Sub

Sub AySy__Tst()
Dim Act$(): Act = AySy(Array(1, 2, 3))
Ass Sz(Act) = 3
Ass Act(0) = 1
Ass Act(1) = 2
Ass Act(2) = 3
End Sub

Private Sub AyWhExclAtCnt__Tst()
Dim A(): A = Array(1, 2, 3, 4, 5)
Dim Act: Act = AyWhExclAtCnt(A, 1, 2)
AyPair_EqChk Array(1, 4, 5), Act
End Sub

Private Sub AyWhExclIxAy__Tst()
Dim A(): A = Array("a", "b", "c", "d", "e", "f")
Dim IxAy: IxAy = Array(1, 3)
Dim Exp: Exp = Array("a", "c", "e", "f")
Dim Act: Act = A: AyWhExclIxAy Act, IxAy
Ass Sz(Act) = 4
Dim J%
For J = 0 To 3
    Ass Act(J) = Exp(J)
Next
End Sub

Private Sub IxAy_IsParitial_of_0toU__Tst()
Ass IxAy_IsParitial_of_0toU(ApLngAy(1, 2, 3, 4), 4) = True
Ass IxAy_IsParitial_of_0toU(ApLngAy(0, 1, 2, 3, 4), 4) = True
Ass IxAy_IsParitial_of_0toU(ApLngAy(1, 1, 3, 4), 4) = False
Ass IxAy_IsParitial_of_0toU(ApLngAy(5, 3, 4), 4) = False
End Sub

Sub LinesEndTrim__Tst()
Dim Lines$: Lines = RplVBar("lksdf|lsdfj|||")
Dim Act$: Act = LinesEndTrim(Lines)
Debug.Print Act & "<"
Stop
End Sub

Private Sub SyTrim__Tst()
AyDmp SyTrim(ApSy(1, 2, 3, "  a"))
End Sub

Private Function AySrtInToIxAy_Ix&(Ix&(), A, V, Des As Boolean)
Dim I, O&
If Des Then
    For Each I In Ix
        If V > A(I) Then AySrtInToIxAy_Ix& = O: Exit Function
        O = O + 1
    Next
    AySrtInToIxAy_Ix& = O
    Exit Function
End If
For Each I In Ix
    If V < A(I) Then AySrtInToIxAy_Ix& = O: Exit Function
    O = O + 1
Next
AySrtInToIxAy_Ix& = O
End Function
