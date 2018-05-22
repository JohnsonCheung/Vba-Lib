Attribute VB_Name = "IdeSrc1"
Option Explicit

Function DclEnmBdyLy(A$(), EnmNm$) As String()
Dim B%: B = DclEnmLx(A, EnmNm): If B = -1 Then Exit Function
Dim O$(), J%
For J = B To UB(A)
   Push O, A(J)
   If HasPfx(A(J), "End Enum") Then DclEnmBdyLy = O: Exit Function
Next
Stop
End Function

Function DclEnmLx%(A$(), EnmNm$)
Dim U%: U = UB(A)
Dim O%, L$
For O = 0 To U
   If SrcLin_IsEmn(A(O)) Then
       L = A(O)
       L = SrcLin_RmvMdy(L)
       L = LinRmvT1(L)
       If LinT1(L) = EnmNm Then
           DclEnmLx = O: Exit Function
       End If
   End If
Next
DclEnmLx = -1
End Function

Function DclEnmNy(A$()) As String()
If AyIsEmp(A) Then Exit Function
Dim I, O$()
For Each I In A
   PushNonEmp O, NewSrcLin(I).EnmNm
Next
DclEnmNy = O
End Function

Function DclHasTy(A$(), TyNm$) As Boolean
If AyIsEmp(A) Then Exit Function
Dim I
For Each I In A
   If HasPfx(I, "Type") Then If LinT2(I) = TyNm Then DclHasTy = True: Exit Function
Next
End Function

Function DclNEnm%(A$())
If AyIsEmp(A) Then Exit Function
Dim I, O%
For Each I In A
   If SrcLin_IsEmn(I) Then O = O + 1
Next
DclNEnm = O
End Function

Function DclTyFmIx&(A$(), TyNm$)
Dim J%, L$
For J = 0 To UB(A)
   If SrcLin_TyNm(A(J)) = TyNm Then DclTyFmIx = J: Exit Function
Next
DclTyFmIx = -1
End Function

Function DclTyFmTo(A$(), TyNm$) As FmtO
Dim FmI&: FmI = DclTyFmIx(A, TyNm)
Dim ToI&: ToI = DclTyToIx(A, FmI)
DclTyFmTo = NewFmTo(FmI, ToI)
End Function

Function DclTyIx%(A$(), TyNm)
Dim I%
   For I = 0 To UB(A)
       If SrcLin_TyNm(A(I)) = TyNm Then
           DclTyIx% = I
           Exit Function
       End If
   Next
DclTyIx = -1
End Function

Function DclTyLines$(A$(), TyNm$)
Stop
DclTyLines = JnCrLf(DclTyLy(A, TyNm))
End Function

Function DclTyLy(A$(), TyNm$) As String()
DclTyLy = AyWhFmTo(A, DclTyFmTo(A, TyNm))
End Function

Function DclTyNy(A$(), Optional TyNmPatn$ = ".") As String()
If AyIsEmp(A) Then Exit Function
Dim O$(), L, M$, Re As RegExp
Set Re = NewRe(TyNmPatn)
For Each L In A
   M = SrcLin_TyNm(L)
   If Re.Test(M) Then
       PushNonEmp O, M
   End If
Next
DclTyNy = O
End Function

Function DclTyToIx(A$(), FmI&)
If 0 > FmI Then DclTyToIx = -1: Exit Function
Dim O&
For O = FmI + 1 To UB(A)
   If HasPfx(A(O), "End Type") Then DclTyToIx = O: Exit Function
Next
DclTyToIx = -1
End Function

Function MthDrs_Ky(A As Drs) As String()
Dim Ty$, Mdy$, MthNm$, K$, IxAy%(), Dr, O$()
IxAy = FnyIxAy(A.Fny, "Mdy MthNm Ty")
If AyIsEmp(A.Dry) Then Exit Function
For Each Dr In A.Dry
    'Debug.Print Mdy, MthNm, Ty
    DrIxAy_Asg Dr, IxAy, Mdy, MthNm, Ty
    K = MthNm & ":" & Ty & ":" & Mdy
    Push O, K
Next
MthDrs_Ky = O
End Function

Function SrcAddMthIfNotExist(A$(), MthNm$, NewMthLy$()) As String()
If SrcHasMth(A, MthNm) Then
   SrcAddMthIfNotExist = A
Else
   SrcAddMthIfNotExist = AyAddAp(A, NewMthLy)
End If
End Function

Function SrcBdyLines$(A$())
SrcBdyLines = JnCrLf(SrcBdyLy(A))
End Function

Function SrcBdyLnoCnt(A$()) As LnoCnt
Dim Lno&
Dim Cnt&
   Lno = SrcDclCnt(A) + 1
   Cnt = Sz(A) - Lno + 1
SrcBdyLnoCnt = NewLnoCnt(Lno, Cnt)
End Function

Function SrcBdyLy(A$()) As String()
SrcBdyLy = AyWhFm(A, SrcDclCnt(A))
End Function

Function SrcCmpLy(A1$(), A2$()) As String()
Dim D1 As Dictionary: Set D1 = SrcDic(A1)
Dim D2 As Dictionary: Set D2 = SrcDic(A2)
Dim Rslt As DCRslt: Rslt = DicCmpRslt(D1, D2)
SrcCmpLy = DCRsltLy(Rslt, "Bef-Srt", "Aft-Srt")
End Function

Function SrcContLin$(A$(), FmIx&)
If FmIx = -1 Then Exit Function
Const CSub$ = "SrcContLinFm"
Dim J&, I$
Dim O$, IsCont As Boolean
For J = FmIx To UB(A)
   I = A(J)
   O = O & LTrim(I)
   IsCont = HasSfx(O, " _")
   If IsCont Then O = RmvSfx(O, " _")
   If Not IsCont Then Exit For
Next
If IsCont Then Er CSub, "each lines {Src} ends with sfx _, which is impossible"
SrcContLin = O
End Function

Function SrcDcl(A$()) As String()
If AyIsEmp(A) Then Exit Function
Dim I&
   I = SrcDclToLx(A)
If I = -1 Then Exit Function
SrcDcl = AyFstUEle(A, I)
End Function

Function SrcDclCnt&(A$())
Dim I&: I = SrcDclToLx(A)
If I = -1 Then SrcDclCnt = Sz(A): Exit Function
SrcDclCnt = I + 1
End Function

Function SrcDclLines$(A$())
SrcDclLines = JnCrLf(SrcDcl(A))
End Function

Function SrcDclToLx&(A$())
Dim I&
    I = SrcFstMthLx(A)
    If I = -1 Then
        I = UB(A) + 1
    Else
        I = SrcMthLx_MthRmkLx(A, I)
    End If
Dim O&
    For I = I - 1 To 0 Step -1
         If SrcLin_IsCd(A(I)) Then O = I: GoTo X
    Next
    O = -1
X:
SrcDclToLx = O
End Function

Function SrcDic(A$()) As Dictionary
Dim Drs As Drs: Drs = SrcMthDrs(A, WithBdyLines:=True)
Dim Ky$(): Ky = MthDrs_Ky(Drs)
Dim BdyLinesAy$(): BdyLinesAy = DrsStrCol(Drs, "BdyLines")
Dim O As New Dictionary: Set O = AyPair_Dic(Ky, BdyLinesAy)
O.Add "*Dcl", SrcDclLines(A)
Set SrcDic = O
End Function

Function SrcDisMthNy(A$()) As String()
Dim O$(), I
If AyIsEmp(A) Then Exit Function
For Each I In A
   PushNonEmp O, SrcLin_MthNm(I)
Next
SrcDisMthNy = O
End Function

Function SrcEnsMth(T$(), MthNm$, NewMthLy$()) As String()
SrcEnsMth = SrcAddMthIfNotExist(T, MthNm, NewMthLy)
End Function

Function SrcEnsOptCmpDb(A$()) As String()
SrcEnsOptCmpDb = SrcEnsOptXXX(A, "Compare Database")
End Function

Function SrcEnsOptExplicit(A$()) As String()
SrcEnsOptExplicit = SrcEnsOptXXX(A, "Explicit")
End Function

Function SrcEnsOptXXX(A$(), OptXXX$) As String()
If SrcHasOptXXX(A, OptXXX) Then
   SrcEnsOptXXX = A
   Debug.Print "Src (* With Option Explicit *)"
Else
   Debug.Print "Src <-------------------- No Option " & OptXXX
   SrcEnsOptXXX = AyIns(A, "Option " & OptXXX)
End If
End Function

Function SrcEnsSrcItm(T$(), A As SrcItm) As String()
Select Case A.SrcTy
Case eSrcTy.eDtaTy: SrcEnsSrcItm = SrcEnsMth(T, A.Nm, A.Ly)
Case eSrcTy.eMth: SrcEnsSrcItm = SrcRplTy(T, A.Nm, A.Ly)
Case Else: Stop
End Select
End Function

Function SrcEnsSrcItmAy(T$(), A() As SrcItm) As String()
Dim O$(): O = T
Dim J%
For J = 0 To SrcItmUB(A)
   O = SrcEnsSrcItm(O, A(J))
Next
SrcEnsSrcItmAy = O
End Function

Function SrcFstMthLx&(A$())
Dim J%
For J = 0 To UB(A)
   If SrcLin_IsMth(A(J)) Then
       SrcFstMthLx = J
       Exit Function
   End If
Next
SrcFstMthLx = -1
End Function

Function SrcHasMth(A$(), MthNm$) As Boolean
SrcHasMth = SrcMth_Lx(A, MthNm) >= 0
End Function

Function SrcHasOptXXX(A$(), OptXXX$) As Boolean
Dim Ay$()
   Ay = SrcDcl(A)
If AyIsEmp(Ay) Then Exit Function
Dim I
For Each I In Ay
   If I = "Option " & OptXXX Then SrcHasOptXXX = True: Exit Function
Next
End Function

Function SrcLin_MthNmPos%(Lin$)
End Function

Function SrcMthCnt%(A$())
If AyIsEmp(A) Then Exit Function
Dim I, O%
For Each I In A
   If SrcLin_IsMth(I) Then O = O + 1
Next
SrcMthCnt = O
End Function

Function SrcMthDrs(A$(), Optional MdNm$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
SrcMthDrs.Dry = SrcMthDry(A, MdNm, WithBdyLy, WithBdyLines)
SrcMthDrs.Fny = FnyOfMthDrs(WithBdyLy, WithBdyLines)
End Function

Function SrcMthDry(A$(), Optional MdNm$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim MthLxAy&(): MthLxAy = SrcMthLxAy(A)
If AyIsEmp(MthLxAy) Then Exit Function
Dim O()
   Dim Dr()
   Dim MthLx
   Dim BdyLy$()
   For Each MthLx In MthLxAy
       Dr = SrcLin_MthDr(A(MthLx), MthLx + 1, MdNm)
       If WithBdyLy Or WithBdyLines Then
           BdyLy = SrcMthLx_BdyLy(A, MthLx)
           If WithBdyLy Then Push Dr, BdyLy
           If WithBdyLines Then Push Dr, JnCrLf(BdyLy)
       End If
       Push O, Dr
   Next
SrcMthDry = O
End Function

Function SrcMthLin$(A$(), MthNm$)
SrcMthLin = SrcContLin(A, SrcMth_Lx(A, MthNm))
End Function

Function SrcMthLinAy(A$()) As String()
If AyIsEmp(A) Then Exit Function
Dim O$(), I, J&
For Each I In A
   If SrcLin_IsMth(I) Then
       Push O, SrcContLin(A, J)
   End If
   J = J + 1
Next
SrcMthLinAy = O
End Function

Function SrcMthLxAy(A$()) As Long()
If AyIsEmp(A) Then Exit Function
Dim O&(), I, J&
   For Each I In A
       If SrcLin_IsMth(I) Then Push O, J
       J = J + 1
   Next
SrcMthLxAy = O
End Function

Function SrcMthLx_BdyLy(A$(), MthLx) As String()
Dim ToLx%: ToLx = SrcMthLx_ToLx(A, MthLx)
Dim FmLx%: FmLx = SrcMthLx_MthRmkLx(A, MthLx)
Dim FT As FmtO
With FT
   .FmIx = FmLx
   .ToIx = ToLx
End With
Dim O$()
   O = AyWhFmTo(A, FT)
SrcMthLx_BdyLy = O
If AyLasEle(O) = "" Then Stop
End Function

Function SrcMthLx_MthRmkLx&(A$(), MthLx)
Dim M1&
    Dim J&
    For J = MthLx - 1 To 0 Step -1
        If SrcLin_IsCd(A(J)) Then
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
SrcMthLx_MthRmkLx = M2
End Function

Function SrcMthNy(A$(), Optional MthNmPatn$ = ".") As String()
If AyIsEmp(A) Then Exit Function
Dim O$(), L, M$
For Each L In A
   M = SrcLin_MthNm(L)
   If ReTst(M, MthNmPatn) Then
       PushNonEmp O, M
   End If
Next
SrcMthNy = O
End Function

Function SrcMth_BdyLines$(A$(), MthNm$)
SrcMth_BdyLines = JnCrLf(SrcMth_BdyLy(A, MthNm))
End Function

Function SrcMth_BdyLy(A$(), MthNm$) As String()
Dim FmtO() As FmtO: FmtO = SrcMth_FmToAy(A, MthNm)
Dim O$(), J%
For J = 0 To FmTo_UB(FmtO)
   PushAy O, AyWhFmTo(A, FmtO(J))
Next
SrcMth_BdyLy = O
End Function

Function SrcMth_FmToAy(A$(), MthNm$) As FmtO()
Dim IxAy&(), O() As FmtO, M As FmtO
IxAy = SrcMth_LxAy(A, MthNm)
Dim J%
For J = 0 To UB(IxAy)
   M.FmIx = IxAy(J)
   M.ToIx = SrcMthLx_ToLx(A, M.FmIx)
   FmTo_Push O, M
Next
SrcMth_FmToAy = O
End Function

Function SrcMth_Lno%(A$(), MthNm, Optional PrpTy$)
If AyIsEmp(A) Then Exit Function
If PrpTy <> "" Then
   If Not AyHas(Array("Get Let Set"), PrpTy) Then Stop
End If
Dim FunTy$: FunTy = "Property " & PrpTy
Dim Lno&
Lno = 0
Const IMthNm% = 2
Dim M As MthBrk
Dim Lin
For Each Lin In A
   Lno = Lno + 1
   M = SrcLin_MthBrk(Lin)
   If M.MthNm = "" Then GoTo Nxt
   If M.MthNm <> MthNm Then GoTo Nxt
   If PrpTy <> "" Then
       If M.Ty <> FunTy Then GoTo Nxt
   End If
   SrcMth_Lno = Lno
   Exit Function
Nxt:
Next
SrcMth_Lno = 0
End Function

Function SrcMth_LnoCnt(A$(), MthNm$) As LnoCnt
End Function

Function SrcMth_LnoCntAy(A$(), MthNm$) As LnoCnt()
Dim FmAy&(): FmAy = SrcMth_LxAy(A, MthNm)
Dim O() As LnoCnt, J%
Dim ToIx&
Dim FmtO As FmtO
Dim LnoCnt As LnoCnt
For J = 0 To UB(FmAy)
   ToIx = SrcMthLx_ToLx(A, FmAy(J))
   FmtO = NewFmTo(FmAy(J), ToIx)
   LnoCnt = FmTo_LnoCnt(FmtO)
   LnoCnt_Push O, LnoCnt
Next
SrcMth_LnoCntAy = O
End Function

Function SrcMth_Lx%(A$(), MthNm$, Optional Fm&)
Dim I%, Nm$
   For I = Fm To UB(A)
       Nm = SrcLin_MthNm(A(I))
       'If HasPfx(A(I), "Property") Then Stop
       'If Nm <> "" Then Debug.Print Nm
       If Nm = MthNm Then
           SrcMth_Lx% = I
           Exit Function
       End If
   Next
SrcMth_Lx = -1
End Function

Function SrcMth_LxAy(A$(), MthNm$) As Long()
Dim Ix&
   Ix = SrcMth_Lx(A, MthNm)
   If Ix = -1 Then Exit Function

Dim O&()
   Push O, Ix
   If Not HasPfx(SrcLin_MthTy(A(Ix)), "Property") Then
       SrcMth_LxAy = O
       Exit Function
   End If

   Dim J%, Fm&
   For J = 1 To 2
       Fm = Ix + 1
       Ix = SrcMth_Lx(A, MthNm, Fm)
       If Ix = -1 Then
           SrcMth_LxAy = O
           Exit Function
       End If
       Push O, Ix
   Next
SrcMth_LxAy = O
End Function

Function SrcMth_RRCC(A$(), MthNm$) As RRCC
Dim R&, C&, Ix&
Ix = SrcMth_Lx(A, MthNm)
R = Ix + 1
C = SrcLin_MthNmPos(A(Ix))
SrcMth_RRCC = NewRRCC(R, R, C, C + Len(MthNm))
End Function

Function SrcNDisMth%(A$())
SrcNDisMth = Sz(SrcDisMthNy(A))
End Function

Function SrcNMth%(A$())
Dim O%, I
If AyIsEmp(A) Then Exit Function
For Each I In A
   If SrcLin_IsMth(I) Then O = O + 1
Next
SrcNMth = O
End Function

Function SrcNTy%(A$())
If AyIsEmp(A) Then Exit Function
Dim I, O%
For Each I In A
   If SrcLin_IsTy(I) Then O = O + 1
Next
SrcNTy = O
End Function

Function SrcPrvMthNy(A$(), Optional MthNmPatn$ = ".") As String()
If AyIsEmp(A) Then Exit Function
Dim O$(), L
For Each L In A
   With SrcLin_MthBrk(L)
       If .Mdy = "Private" Then
           If .MthNm <> "" Then
               If ReTst(.MthNm, MthNmPatn) Then
                   Push O, .MthNm
               End If
           End If
       End If
   End With
Next
SrcPrvMthNy = O
End Function

Function SrcRmvMth(A$(), MthNm$) As String()
Dim FmToAy() As FmtO
   FmToAy = SrcMth_FmToAy(A, MthNm)
Dim O$()
   O = A
   Dim J%
   For J = FmTo_UB(FmToAy) To 0 Step -1
       O = AyRmvFmTo(O, FmToAy(J))
   Next
SrcRmvMth = O
End Function

Function SrcRmvTy(A$(), TyNm$) As String()
SrcRmvTy = AyRmvFmTo(A, DclTyFmTo(A, TyNm))
End Function

Function SrcRplMth(A$(), MthNm$, NewMthLy$()) As String()
Dim OldMthLines$
   OldMthLines = SrcMth_BdyLines(A, MthNm)
Dim NewMthLines$
   NewMthLines = JnCrLf(NewMthLy)
If OldMthLines = NewMthLines Then
   SrcRplMth = A
   Exit Function
End If
Dim O$()
   O = SrcRmvMth(A, MthNm)
   PushAy O, NewMthLy
SrcRplMth = O

End Function

Function SrcRplTy(A$(), TyNm$, NewTyLy$()) As String()
Dim Dcl$()
Dim FmtO As FmtO
Dim Old$()
   Dcl = SrcDcl(A)
   FmtO = DclTyFmTo(Dcl, TyNm)
   Old = AyWhFmTo(Dcl, FmtO)
If AyIsEq(Old, NewTyLy) Then
   SrcRplTy = A
Else
   SrcRplTy = AyRpl(A, FmtO, NewTyLy)
End If
End Function

Sub ZZ_SrcDic()
LinesDic_Brw SrcDic(ZZSrc)
End Sub

Sub ZZ_SrcMthDrs()
Dim Src$(): Src = MdSrc(IdeMd.Md("ThisWorkbook"))
DrsDmp SrcMthDrs(Src, WithBdyLy:=True)
End Sub

Sub ZZ_SrcMthDry()
Dim Src$(): Src = MdSrc(IdeMd.Md("ThisWorkbook"))
DryDmp SrcMthDry(Src, "IdeSrc")
End Sub

Private Function SrcBdyIx%(A$(), FstMthLx&)
Dim J%
For J = FstMthLx - 1 To 0 Step -1
   If SrcLin_IsCd(A(J)) Then SrcBdyIx = J + 1: Exit Function
Next
SrcBdyIx = 0
End Function

Private Function SrcMthLx_ToLx&(A$(), MthLx)
Const CSub$ = "SrcMthLx_ToLx"
Dim Lin$
   Lin = A(MthLx)

Dim Pfx$
   Pfx = SrcLin_EndLinPfx(Lin)
Dim O&
   For O = MthLx + 1 To UB(A)
       If HasPfx(A(O), Pfx) Then SrcMthLx_ToLx = O: Exit Function
   Next
Er CSub, "{Src}-{MthFmIx} is {MthLin} which does have {FunEndLinPfx} in lines after [MthFmIx]", A, MthLx, Lin, Pfx
End Function

Private Function ZZSrc() As String()
ZZSrc = MdSrc(IdeMd.Md("IdeSrc"))
End Function

Private Sub ZZ_SrcContLin()
Dim O$(3)
O(0) = "A _"
O(1) = "  B _"
O(2) = "C"
O(3) = "D"
Dim Act$: Act = SrcContLin(O, 0)
Ass Act = "A B C"
End Sub

Private Sub ZZ_SrcDcl()
StrBrw SrcDcl(ZZSrc)
End Sub

Private Sub ZZ_SrcDclCnt()
Dim Act%
   Act = SrcDclCnt(ZZSrc)
Ass Act = 2
End Sub

Private Sub ZZ_SrcFstMthLx()
Dim Act%
Act = SrcFstMthLx(ZZSrc)
Ass Act = 2
End Sub

Private Sub ZZ_SrcMthLxAy()
Dim Src$(): Src = ZZSrc
Dim LxAy&(): LxAy = SrcMthLxAy(ZZSrc)
Dim Ay$(): Ay = AyWh_ByIxAy(Src, LxAy)
Dim Dry(): Dry = AyZip(LxAy, Ay)
Dim O$()
O = DrsLy(NewDrs("Lx Lin", AyZip(LxAy, Ay)))
PushAy O, DrsLy(AyDrs(Src))
AyBrw O
End Sub

Private Sub ZZ_SrcMthLx_MthRmkLx()
Dim ODry()
    Dim Src$(): Src = MdSrc(IdeMd.Md("IdeSrcLin"))
    Dim Dr(), Lx&
    Dim J%, IsMth$, RmkLx$, Lin
    For Each Lin In Src
        IsMth = ""
        RmkLx = ""
        If SrcLin_IsMth(Lin) Then
            If Lx = 482 Then Stop
            IsMth = "*Mth"
            RmkLx = SrcMthLx_MthRmkLx(Src, Lx)
            
        End If
        Dr = Array(IsMth, RmkLx, Lin)
        Push ODry, Dr
        Lx = Lx + 1
    Next
Dim Drs As Drs
    Drs = NewDrs("Mth RmkLx Lin", ODry)
DrsBrw Drs
End Sub

Private Sub DclTyLines__Tst()
Debug.Print DclTyLines(MdDcl(CurMd), "AA")
End Sub

Private Sub SrcContLin__Tst()
ZZ_SrcContLin
End Sub

Private Sub SrcDclCnt__Tst()
ZZ_SrcDclCnt
End Sub

Private Sub SrcDcl__Tst()
ZZ_SrcDcl
End Sub

Sub SrcDic__Tst()
Dim Act() As S1S2
Act = DicS1S2Ay(SrcDic(ZZSrc))
AyBrw S1S2Ay_FmtLy(Act)
End Sub

Private Sub SrcFstMthLx__Tst()
ZZ_SrcFstMthLx
End Sub

Private Sub SrcMthDrs__Tst()
DrsBrw SrcMthDrs(ZZSrc)
End Sub

Sub SrcMthDry__Tst()
DryBrw SrcMthDry(ZZSrc, "IdeSrc")
End Sub

Sub SrcMthLxAy__Tst()
Dim Src$(): Src = MdSrc(IdeMd.Md("DaoDb"))
Dim Ay$(): Ay = AyWhIxAy(Src, SrcMthLxAy(Src))
AyBrw Ay
End Sub

Private Sub SrcMthNy__Tst()
Dim Act$()
   Act = SrcMthNy(ZZSrc)
   AyBrw Act
End Sub

Sub SrcMth_BdyLy__Tst()
Dim Src$(): Src = ZZSrc
Dim MthNm$: MthNm = "A"
Dim Act$()
Act = SrcMth_BdyLy(Src, MthNm)
End Sub