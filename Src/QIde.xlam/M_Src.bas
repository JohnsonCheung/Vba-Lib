Attribute VB_Name = "M_Src"
Option Explicit

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
   Lno = SrcDclLinCnt(A) + 1
   Cnt = Sz(A) - Lno + 1
Set SrcBdyLnoCnt = LnoCnt(Lno, Cnt)
End Function

Function SrcBdyLy(A$()) As String()
SrcBdyLy = AyWhFm(A, SrcDclLinCnt(A))
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

Function SrcDclLinCnt%(A$())
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
SrcDclLinCnt = O
End Function

Function SrcDclLines$(A$())
SrcDclLines = JnCrLf(SrcDclLy(A))
End Function

Function SrcDclLy(A$()) As String()
If AyIsEmp(A) Then Exit Function
Dim N&
   N = SrcDclLinCnt(A)
If N = 0 Then Exit Function
SrcDclLy = AyFstNEle(A, N)
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
   Ay = SrcDclLy(A)
If AyIsEmp(Ay) Then Exit Function
Dim I
For Each I In Ay
   If I = "Option " & OptXXX Then SrcHasOptXXX = True: Exit Function
Next
End Function


Function SrcMthCnt%(A$())
If AyIsEmp(A) Then Exit Function
Dim I, O%
For Each I In A
   If SrcLin_IsMth(I) Then O = O + 1
Next
SrcMthCnt = O
End Function

Function SrcMthDrs(A$(), Optional MdNm$, Optional MdTy$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
Dim Dry(): Dry = SrcMthDry(A, MdNm, MdTy, WithBdyLy, WithBdyLines)
Dim Fny$(): Fny = FnyOfMthDrs(WithBdyLy, WithBdyLines)
Set SrcMthDrs = Drs(Fny, Dry)
End Function

Function SrcMthDry(A$(), Optional MdNm$, Optional MdTy$, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim MthLxAy&(): MthLxAy = SrcMthLxAy(A)
If AyIsEmp(MthLxAy) Then Exit Function
Dim O()
   Dim Dr()
   Dim MthLx
   Dim BdyLy$()
   For Each MthLx In MthLxAy
       Dr = SrcLin_MthDr(A(MthLx), MthLx + 1, MdNm, MdTy)
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



Function SrcMthNy(A$(), Optional MthNmPatn$ = ".") As String()
Stop '
'If AyIsEmp(A) Then Exit Function
'Dim O$(), L, M$
''Dim R As Re: ' Set R = Re(MthNmPatn)
'For Each L In A
'   M = SrcLin_MthNm(L)
'   If R.Tst(M) Then
'       PushNonEmp O, M
'   End If
'Next
'SrcMthNy = O
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
Dim O$(), L, R As RegExp
Set R = Re(MthNmPatn)
For Each L In A
   With SrcLin_MthBrk(L)
       If .Mdy = "Private" Then
           If .MthNm <> "" Then
               If R.Test(.MthNm) Then
                   Push O, .MthNm
               End If
           End If
       End If
   End With
Next
SrcPrvMthNy = O
End Function

Function SrcRmvMth(A$(), MthNm$) As String()
Dim FmToAy() As FmTo
   FmToAy = SrcMth_FmToAy(A, MthNm)
Dim O$()
   O = A
   Dim J%
   For J = UB(FmToAy) To 0 Step -1
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
Dim FmTo As FmTo
Dim Old$()
   Dcl = SrcDclLy(A)
   Set FmTo = DclTyFmTo(Dcl, TyNm)
   Old = AyWhFmTo(Dcl, FmTo)
If AyIsEq(Old, NewTyLy) Then
   SrcRplTy = A
Else
   SrcRplTy = AyRpl(A, FmTo, NewTyLy)
End If
End Function

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
ZZSrc = MdSrc(Md("IdeSrc"))
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
StrBrw SrcDclLy(ZZSrc)
End Sub

Private Sub ZZ_SrcDclLinCnt()
Dim Act%
   Act = SrcDclLinCnt(ZZSrc)
Ass Act = 2
End Sub

Private Sub ZZ_SrcDic()
Stop '
'Dim Act() As S1S2
'Act = Dix(SrcDic(ZZSrc)).S1S2Ay
'AyBrw S1S2Ay_FmtLy(Act)
''LinesDic_Brw SrcDic(ZZSrc)
End Sub

Private Sub ZZ_SrcFstMthLx()
Dim Act%
Act = SrcFstMthLx(ZZSrc)
Ass Act = 2
End Sub

Private Sub ZZ_SrcMthDrs()
'DrsBrw SrcMthDrs(ZZSrc)
Dim Src$(): Src = MdSrc(M_Md.Md("ThisWorkbook"))
DrsDmp SrcMthDrs(Src, WithBdyLy:=True)
End Sub

Private Sub ZZ_SrcMthDry()
Dim Src$(): Src = MdSrc(Md("ThisWorkbook"))
DryDmp SrcMthDry(Src, "IdeSrc")
End Sub

Private Sub ZZ_SrcMthDry1()
DryBrw SrcMthDry(ZZSrc, "IdeSrc")
End Sub

Private Sub ZZ_SrcMthLxAy()
Stop '
'Dim Src$(): Src = ZZSrc
'Dim LxAy&(): LxAy = SrcMthLxAy(ZZSrc)
'Dim Ay$(): Ay = AyWh_ByIxAy(Src, LxAy)
'Dim Dry(): Dry = AyZip(LxAy, Ay)
'Dim O$()
'O = DrsLy(Drs("Lx Lin", AyZip(LxAy, Ay)))
'PushAy O, DrsLy(AyDrs(Src))
'AyBrw O
End Sub

Private Sub ZZ_SrcMthLxAy1()
Dim Src$(): Src = MdSrc(Md("DaoDb"))
Dim Ay$(): Ay = AyWhIxAy(Src, SrcMthLxAy(Src))
aybrw Ay
End Sub


Private Sub ZZ_SrcMthNy()
Dim Act$()
   Act = SrcMthNy(ZZSrc)
   aybrw Act
End Sub

Function SrcItmLy(A As SrcItm) As String()
SrcItmLy = A.Ly
End Function
Function SrcItmSz%(A() As SrcItm)
On Error Resume Next
SrcItmSz = UBound(A) + 1
End Function
Function SrcItmUB%(A() As SrcItm)
SrcItmUB = SrcItmSz(A) - 1
End Function
Sub SrcItmAyDmp(A() As SrcItm)
Dim J%
For J = 0 To SrcItmUB(A)
AyDmp SrcItmLy(A(J))
Next
End Sub
Sub SrcItmDmp(A As SrcItm)
AyDmp SrcItmLy(A)
End Sub
Sub SrcItmPush(O() As SrcItm, M As SrcItm)
Dim N%: N = SrcItmSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub SrcItmPushAy(O() As SrcItm, M() As SrcItm)
Dim J%
For J = 0 To SrcItmUB(M)
   SrcItmPush O, M(J)
Next
End Sub
Function SrcLinInfFny() As String()
Static X As Boolean, Y$()
If Not X Then
    X = True
    Y = SplitSpc("Md Lno Lin EnmNm IsBlank IsEmn IsMth IsPrpLin IsRmk IsTy Mdy MthNm MthTy NoMdy PrpTy TyNm")
End If
SrcLinInfFny = Y
End Function
