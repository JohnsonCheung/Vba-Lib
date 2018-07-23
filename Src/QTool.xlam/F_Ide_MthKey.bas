Attribute VB_Name = "F_Ide_MthKey"
Option Explicit
'Mth2DLin = MthFmLno.MthLinCnt.MthLin
'Mth4DLin = Pj.Md.MthFmLno.MthLinCnt.MthLin
'MthDr = { Pj MdTy Md Pfx Nm Lno Cnt ShtTy ShtMdy RetTy Prm Rmk [Lines] MthCnt [SamLinesCnt]} Sz = 14 or 15
Sub ZZ_FnyOf_MthDr()
AyDmp FnyOf_MthDr(True)
End Sub
Function FnyOf_MthDr(InclMthLines As Boolean) As String()
Const X$ = "Pj MdTy Md MthPfx Nm Lno Cnt MthShtTy Mdy RetTy Prm Rmk? MthCnt?"
Dim A$, B$, C$
If InclMthLines Then A = " Lines": B = " SamLinesCnt"
C = FmtQQ(X, A, B)
Debug.Print C
Debug.Print X
FnyOf_MthDr = SslSy(C)
End Function
 
Function FxOf_Mth$()
FxOf_Mth = CurPjPth & "VbeMthKey.xlsx"
End Function

Function MdMth4DLinAy(A As CodeModule, InclMthLines As Boolean) As String()
MdMth4DLinAy = AyAddPfx(SrcMth2DLinAy(MdSrc(A), InclMthLines), MdDNm(A) & ".")
End Function

Function MdMthWs(A As CodeModule) As Worksheet
Set MdMthWs = WsVis(SqWs(MdMthSq(A)))
End Function

Function PjMthWs(A As CodeModule) As Worksheet
Set PjMthWs = WsVis(SqWs(PjMthSq(A)))
End Function
Function Mth4DLinAy_Sq(A$(), InclMthLines As Boolean) As Variant()
Mth4DLinAy_Sq = DryFny_Sq(Mth4DLinAy_Dry(A, InclMthLines), FnyOf_MthDr(InclMthLines))
End Function

Function Mth4DLinAy_Dry(A$(), InclMthLines As Boolean) As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), J%, U&
U = UB(A)
ReDim O(U)
For J = 0 To U
    O(J) = Mth4DLin_MthDr(A(J), InclMthLines)
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

Function Mth4DLin_MthDr(A$, InclMthLines As Boolean) As Variant()
Dim L$: L = A
Dim Pj$: Pj = LinShiftDTerm(L)
Dim Md$: Md = LinShiftDTerm(L)
Dim Lno%: Lno = LinShiftDTerm(L)
Dim Cnt%: Cnt = LinShiftDTerm(L)
Dim MthLines$
    If InclMthLines Then
        With Brk(L, "{|}")
            L = .S1
            MthLines = .S2
        End With
    End If
If Cnt <= 0 Then Stop
Mth4DLin_MthDr = MthLin_MthDr(L, Pj, Md, Lno, Cnt, MthLines)
End Function

Function MthDr(A As Mth, Optional X As FmToLno, Optional InclMthLines As Boolean) As Variant()
Dim Lno%, Cnt%
If IsNothing(X) Then Set X = MthFmToLno(A)
Dim Lines$: If InclMthLines Then Lines = MthLines(A)
MthDr = MthLin_MthDr(MthLin(A), MthPjNm(A), MthMdNm(A), X.FmLno, FmToLno_LinCnt(X), Lines)
End Function

Function MdMthKy(A As CodeModule, Optional IsWrap As Boolean) As String()
Dim PjN$: PjN = MdPjNm(A)
Dim MdN$: MdN = MdNm(A)
MdMthKy = SrcMthKy(MdSrc(A), PjN, MdN, IsWrap)
End Function

Function SrcMthKy(A$(), Optional PjNm$ = "Pj", Optional MdNm$ = "Md", Optional IsWrap As Boolean) As String()
Dim L$(): L = SrcMthLinAy(A) ' MthLinAy
SrcMthKy = AyMapAvSy(L, "MthLin_MthKey", Array(PjNm, MdNm, IsWrap))
End Function

Function MdMthDrs(A As CodeModule) As Drs
Set MdMthDrs = Drs(SplitSsl(""), MdMthDry(A))
End Function

Function MdMthDry(A As CodeModule) As Variant()
Dim O()
MdMthDry = O
End Function

Function MthLin_MthDr(A, PjNm$, MdNm$, MthLno%, MthLinCnt%, MthLines$) As Variant()
Dim MdTy$, MthPfx$, Mth$, MthShtTy$, Mdy$, RetTy$, Prm$, Rmk$
MdTy = MdTyNm(Md(PjNm & "." & MdNm))
Dim L$: L = A
Mdy = LinShiftMdy(L)
MthShtTy = MthTy_MthShtTy(LinShiftMthTy(L))
Mth = LinShiftNm(L)
RetTy = LinShiftTySfxChr(L)
Prm = LinShiftBktEnclosedStr(L): If Prm = "" Then Stop
If RetTy <> "" Then
    If L <> "" Then Stop
Else
    Dim T$: T = LinShiftT1(L)
    If T <> "" Then
        Select Case True
        Case T = "As":
            If RetTy <> "" Then Stop
            RetTy = LinShiftT1(L)
            If L <> "" Then
                If Left(L, 1) <> "'" Then Stop
                Rmk = Mid(L, 2)
            End If
        Case Left(T, 1) = "'"
            Rmk = Mid(L, 2)
        Case Else: Stop
        End Select
    End If
End If
MthPfx = MthNm_MthPfx(Mth)
If MthLines = "" Then
    MthLin_MthDr = Array(PjNm, MdTy, MdNm, MthPfx, Mth, MthLno, MthLinCnt, MthShtTy, Mdy, RetTy, Prm, Rmk)
Else
    MthLin_MthDr = Array(PjNm, MdTy, MdNm, MthPfx, Mth, MthLno, MthLinCnt, MthShtTy, Mdy, RetTy, Prm, Rmk, MthLines)
End If
End Function

Function AyOfAy_Ay(A)
If Sz(A) = 0 Then Exit Function
Dim O: O = A(0)
Dim I, J%
For J = 1 To UB(A)
    PushAy O, A(J)
Next
AyOfAy_Ay = O
End Function

Function PjMth4DLinAy(A As VBProject, InclMthLines As Boolean) As String()
'Mth4DLin is: Pj.Md.MthLinCnt.MthLin
PjMth4DLinAy = AyOfAy_Ay(AyMapXP(PjMbrAy(A), "MdMth4DLinAy", InclMthLines))
End Function

Function PjMthKyWs(A As VBProject) As Worksheet
Set PjMthKyWs = WsVis(SqWs(PjMthKySq(A)))
End Function

Sub ZZ_SrcMth2DLinAy()
Dim A$(): A = SrcMth2DLinAy(ZZSrc, True)
End Sub

Function ZZSrc() As String()
ZZSrc = MdSrc(Md("QTool.F_Ide_MthKey"))
End Function

Function SrcMth2DLinAy(A$(), InclMthLines As Boolean) As String()
'MthDLin is: MthFmLno.MthLinCnt.MthLin{|}MthLines
'        or: MthFmLno.MthLinCnt.MthLin
'        depends on InclMthLines
Dim F% ' MthFmLno
Dim C% ' MthLinCnt
Dim L() As FmToLx: L = SrcAllMthFmToLxAy(A)
If Sz(L) = 0 Then Exit Function
Dim O$(), LLL As FmToLx, LL
For Each LL In L
    Set LLL = CvFmToLx(LL)
    F = LLL.FmLx + 1
    C = FmToLx_LinCnt(LLL)
    If InclMthLines Then
        Dim MthLines$: MthLines = JnCrLf(AyWhFmTo(A, LLL.FmLx, LLL.ToLx))
        Push O, F & "." & C & "." & SrcContLin(A, LLL.FmLx) & "{|}" & MthLines
    Else
        Push O, F & "." & C & "." & SrcContLin(A, LLL.FmLx)
    End If
Next
SrcMth2DLinAy = O
End Function

Function VbeMth4DLinAy(A As Vbe, InclMthLines As Boolean) As String()
VbeMth4DLinAy = AyOfAy_Ay(AyMapXP(VbePjAy(A), "PjMth4DLinAy", InclMthLines))
End Function

Function VbeMthSq(A As Vbe, InclMthLines As Boolean) As Variant()
Dim B$(): B = VbeMth4DLinAy(A, InclMthLines)
VbeMthSq = Mth4DLinAy_Sq(B, InclMthLines)
End Function

Function VbeMthWb(A As Vbe, InclMthLines As Boolean) As Workbook
Dim O As Workbook: Set O = FxWb(FxOf_Mth)
Dim Ws As Worksheet: Set Ws = O.Sheets("Data")
Ws.Cells.Delete
Dim A1 As Range, Sq()
Set A1 = WsA1(Ws)
Sq = VbeMthSq(A, InclMthLines)
CellPutSq A1, Sq, "Data"
WsVis Ws
WbRfh O
Set VbeMthWb = O
End Function

Function VbeMthWs(A As Vbe, InclMthLines As Boolean) As Worksheet
Set VbeMthWs = WsVis(SqWs(VbeMthSq(A, InclMthLines)))
End Function

Function WbOf_Mth() As Workbook
Set WbOf_Mth = FxWb(FxOf_Mth)
End Function

Sub AAA()
ZZ_VbeMthWb
End Sub

Sub ZZ_MthDr()
Dim M As Mth: Set M = Mth(Md("F_Ide_MthKey"), "VbeMthWb")
Dim A$(): A = FnyOf_MthDr(InclMthLines:=True)
Dim B():  B = MthDr(M, InclMthLines:=True)
AyDmp AyAB_FmtLy(A, B)
End Sub

Sub ZZ_VbeMthWb()
WbVis VbeMthWb(CurVbe, True)
End Sub

Sub ZZ_VbeMthWs()
WsVis VbeMthWs(CurVbe, True)
End Sub
