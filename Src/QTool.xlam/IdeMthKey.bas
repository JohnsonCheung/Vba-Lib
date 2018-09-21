Attribute VB_Name = "IdeMthKey"
Option Explicit

Function AyOfAy_Ay(A)
If Sz(A) = 0 Then Exit Function
Dim O: O = A(0)
Dim I, J%
For J = 1 To UB(A)
    PushAy O, A(J)
Next
AyOfAy_Ay = O
End Function

Function CvItr(A) As Collection
Set CvItr = A
End Function

Function FnyOf_MthDr(InclMthLines As Boolean) As String()
Const X$ = "Pj MdTy Md MthPfx Nm Lno Cnt MthShtTy Mdy RetTy Prm Rmk? MthCnt?"
Dim A$, B$, C$
If InclMthLines Then A = " Lines": B = " SamLinesCnt"
C = FmtQQ(X, A, B)
FnyOf_MthDr = SslSy(C)
End Function

Function VbeMthFx$()
VbeMthFx = CurPjPth & "PjMthKey.xlsx"
End Function

Function IItrItr(A As Collection) As Collection
Dim O As New Collection, I
For Each I In A
    ItrPushItr O, CvItr(I)
Next
Set IItrItr = O
End Function

Sub ItrPushItr(O As Collection, M As Collection)
Dim I
For Each I In M
    O.Add I
Next
End Sub

Function MdMth4DLinAy(A As CodeModule, InclMthLines As Boolean) As String()
MdMth4DLinAy = AyAddPfx(SrcMth2DLinAy(MdSrc(A), InclMthLines), MdDNm(A) & ".")
End Function

Function MdMthDrs(A As CodeModule) As Drs
Set MdMthDrs = Drs(SplitSsl(""), MdMthDry(A))
End Function

Function MdMthDry(A As CodeModule) As Variant()
Dim O()
MdMthDry = O
End Function

Function MdMthKy(A As CodeModule, Optional IsWrap As Boolean) As String()
Dim PjN$: PjN = MdPjNm(A)
Dim MdN$: MdN = MdNm(A)
MdMthKy = SrcMthKy(MdSrc(A), PjN, MdN, IsWrap)
End Function

Function MdMthWs(A As CodeModule) As Worksheet
Set MdMthWs = WsVis(SqWs(MdMthSq(A)))
End Function

Function Mth4DLinAy_Dry(A$(), InclMthLines As Boolean) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), J%, U&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Mth4DotDr(A(J), InclMthLines)
Next
Dim D As Dictionary
Dim ColIx%
Dim Fny$()
    Fny = FnyOf_MthDr(True)
'-- Add MthCnt Col
    ColIx = AyIx(Fny, "Nm")
    Set D = DryCntDic(O, ColIx)
    O = DryAddColByDic(O, ColIx, D)
'-- Add SamLinesCnt Col
If InclMthLines Then
    ColIx = AyIx(Fny, "Lines")
    Set D = DryCntDic(O, ColIx)
    O = DryAddColByDic(O, ColIx, D)
End If
Mth4DLinAy_Dry = O
End Function

Function Mth4DLinAy_Sq(A$(), InclMthLines As Boolean) As Variant()
Mth4DLinAy_Sq = DryFny_Sq(Mth4DLinAy_Dry(A, InclMthLines), FnyOf_MthDr(InclMthLines))
End Function

Function Mth4DotDr(Mth4DLin$, InclMthLines As Boolean) As Variant()
Dim L$: L = Mth4DLin
Dim Pj$: Pj = ShiftDTerm(L)
Dim Md$: Md = ShiftDTerm(L)
Dim Lno%: Lno = ShiftDTerm(L)
Dim Cnt%: Cnt = ShiftDTerm(L)
Dim MthLines$
    If InclMthLines Then
        With Brk(L, "{|}")
            L = .S1
            MthLines = .S2
        End With
    End If
If Cnt <= 0 Then Stop
Mth4DotDr = LinMthDr(L, Pj, Md, Lno, Cnt, MthLines)
End Function

Function MthDr(A As Mth, Optional X As FTNo, Optional InclMthLines As Boolean) As Variant()
Dim Lno%, Cnt%
If IsNothing(X) Then Set X = MthFTNo(A)
Dim Lines$: If InclMthLines Then Lines = MthLines(A)
MthDr = LinMthDr(MthLin(A), MthPjNm(A), MthMdNm(A), X.Fmno, FTNo_LinCnt(X), Lines)
End Function
Function AyAsg(A, ParamArray OAp())
Dim Av(): Av = OAp
Dim J%
For J = 0 To UB(Av)
    OAp(J) = A(J)
Next
End Function
Function LinMthDr(A, PjNm$, MdNm$, MthLno%, MthLinCnt%, MthLines$) As Variant()
Dim MdTy$, Pfx$, Mth$, ShtTy$, Mdy$, RetTy$, Prm$, Rmk$, Brk$(), Ay(), Rest$, Sy$()
MdTy = MdTyNm(Md(PjNm & "." & MdNm))
AyAsg ShiftMthBrk(A), Brk, Rest
If Brk(2) = "" Then Exit Function
AyAsg Brk, Mdy, ShtTy, Mth
AyAsg ShiftTySfxChr(Rest), RetTy, Rest
AyAsg ShiftBktStr(Rest), Prm, Rest
If RetTy <> "" Then
    If Rest <> "" Then Stop
Else
    Dim T$: AyAsg ShiftT1(Rest), T, Rest
    If T <> "" Then
        Select Case True
        Case T = "As":
            If RetTy <> "" Then Stop
            AyAsg ShiftT1(Rest), RetTy, Rest
            If Rest <> "" Then
                If Left(Rest, 1) <> "'" Then Stop
                Rmk = Mid(Rest, 2)
            End If
        Case Left(T, 1) = "'"
            Rmk = Mid(Rest, 2)
        Case Else: Stop
        End Select
    End If
End If
Pfx = MthPfx(Mth)
If MthLines = "" Then
    LinMthDr = Array(PjNm, MdTy, MdNm, Pfx, Mth, MthLno, MthLinCnt, ShtTy, Mdy, RetTy, Prm, Rmk)
Else
    LinMthDr = Array(PjNm, MdTy, MdNm, Pfx, Mth, MthLno, MthLinCnt, ShtTy, Mdy, RetTy, Prm, Rmk, MthLines)
End If
End Function

Function PjMth4DLinAy(A As VBProject, InclMthLines As Boolean) As String()
'Mth4DLin is: Pj.Md.MthLinCnt.MthLin
PjMth4DLinAy = AyOfAy_Ay(AyMapXP(PjCdMdAy(A), "MdMth4DLinAy", InclMthLines))
End Function

Function PjMthKyWs(A As VBProject) As Worksheet
Set PjMthKyWs = WsVis(SqWs(PjMthKySq(A)))
End Function

Function PjMthWs(A As CodeModule) As Worksheet
Set PjMthWs = WsVis(SqWs(PjMthSq(A)))
End Function

Function SrcMth2DLinAy(A$(), InclMthLines As Boolean) As String()
'MthDLin is: MthFmno.MthLinCnt.MthLin{|}MthLines
'        or: MthFmno.MthLinCnt.MthLin
'        depends on InclMthLines
Dim F% ' MthFmno
Dim C% ' MthLinCnt
Dim L() As FTIx: L = SrcMthFTIxAy(A)
If Sz(L) = 0 Then Exit Function
Dim O$(), LLL As FTIx, LL
For Each LL In L
    Set LLL = CvFTIx(LL)
    F = LLL.Fmix + 1
    C = FTIx_LinCnt(LLL)
    If InclMthLines Then
        Dim MthLines$: MthLines = JnCrLf(AyWhFmTo(A, LLL.Fmix, LLL.Toix))
        Push O, F & "." & C & "." & SrcContLin(A, LLL.Fmix) & "{|}" & MthLines
    Else
        Push O, F & "." & C & "." & SrcContLin(A, LLL.Fmix)
    End If
Next
SrcMth2DLinAy = O
End Function

Function SrcMthKy(A$(), Optional PjNm$ = "Pj", Optional MdNm$ = "Md", Optional IsWrap As Boolean) As String()
Dim L$(): L = SrcMthLinAy(A) ' MthLinAy
SrcMthKy = AyMapAvSy(L, "MthLin_MthKey", Array(PjNm, MdNm, IsWrap))
End Function

Function VbeMth4DLinAy(A As Vbe, InclMthLines As Boolean) As String()
VbeMth4DLinAy = AyOfAy_Ay(AyMapXP(VbePjAy(A), "PjMth4DLinAy", InclMthLines))
End Function

Function VbeMthSq(A As Vbe, InclMthLines As Boolean) As Variant()
Dim B$(): B = VbeMth4DLinAy(A, InclMthLines)
VbeMthSq = Mth4DLinAy_Sq(B, InclMthLines)
End Function
Function CurVbeMthWb(Optional InclMthLines As Boolean) As Workbook
Set CurVbeMthWb = VbeMthWb(CurVbe, InclMthLines)
End Function

Function VbeMthWb(A As Vbe, InclMthLines As Boolean) As Workbook
Dim Fx$
Fx = VbeMthFx
If FfnIsExist(Fx) Then Kill Fx
Dim O As Workbook: Set O = NewWb("Data")
O.SaveAs Fx
Dim Ws As Worksheet: Set Ws = O.Sheets("Data")
Ws.Cells.Delete
Dim A1 As Range, Sq()
Set A1 = WsA1(Ws)
Sq = VbeMthSq(A, InclMthLines)
SqLo Sq, A1, "Data"
WsVis Ws
WbRfh O
Set VbeMthWb = O
End Function

Function VbeMthWs(A As Vbe, InclMthLines As Boolean) As Worksheet
Set VbeMthWs = WsVis(SqWs(VbeMthSq(A, InclMthLines)))
End Function

Function WbOf_Mth() As Workbook
Set WbOf_Mth = FxWb(VbeMthFx)
End Function

'Mth2DLin = MthFmno.MthLinCnt.MthLin
'Mth4DLin = Pj.Md.MthFmno.MthLinCnt.MthLin
'MthDr = { Pj MdTy Md Pfx Nm Lno Cnt ShtTy ShtMdy RetTy Prm Rmk [Lines] MthCnt [SamLinesCnt]} Sz = 14 or 15
Private Sub ZZ_FnyOf_MthDr()
D FnyOf_MthDr(True)
End Sub

Private Sub ZZ_MthDr()
Dim M As Mth: Set M = Mth(Md("F_Ide_MthKey"), "VbeMthWb")
Dim A$(): A = FnyOf_MthDr(InclMthLines:=True)
Dim B():  B = MthDr(M, InclMthLines:=True)
D AyabFmt(A, B)
End Sub
Private Sub ZZ_SrcMth2DLinAy()
Dim A$(): A = SrcMth2DLinAy(CurSrc, True)
End Sub

Private Sub ZZ_VbeMthWb()
WbVis VbeMthWb(CurVbe, True)
End Sub

Private Sub ZZ_VbeMthWs()
WsVis VbeMthWs(CurVbe, True)
End Sub
