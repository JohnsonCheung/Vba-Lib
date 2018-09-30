Attribute VB_Name = "IdeMthKey"
Option Explicit
Private Const Hom$ = "C:\Users\user\Desktop\MHD\SAPAccessReports\"
Private Const StkShpRateFb$ = Hom & "StockShipRate\StockShipRate\StockShipRate (ver 1.0).accdb"
Private Const TaxExpCmpFb$ = Hom & "TaxExpCmp\TaxExpCmp\TaxExpCmp v1.3.accdb"
Private Const StkShpCstFb$ = Hom & "StockShipCost\StockShipCost (ver 1.0).accdb"
Private Const TaxRateAlertFb$ = Hom & "TaxRateAlert\TaxRateAlert\TaxRateAlert (ver 1.3).accdb"
Function AppFbAy() As String()
Push AppFbAy, StkShpCstFb
Push AppFbAy, StkShpRateFb
Push AppFbAy, TaxExpCmpFb
Push AppFbAy, TaxRateAlertFb
End Function
Function FbAcs(A, Optional Vis As Boolean) As Access.Application
Dim O As New Access.Application
O.OpenCurrentDatabase A
O.Visible = Vis
Set FbAcs = O
End Function

Function MdMthWs(A As CodeModule) As Worksheet
Set MdMthWs = WsVis(SqWs(MdMthSq(A)))
End Function

Function MdMthDry(A As CodeModule, Optional B As WhMth, Optional C As MthBrkOpt) As Variant()
MdMthDry = DryAddCC(SrcMthDry(MdBdyLy(A), B, C), MdPjNm(A), MdNm(A))
End Function

Sub Z_MdMthDry()
Brw DryFmtss(MdMthDry(CurMd))
End Sub

Sub Z_VbeMthDry()
Brw DryFmtss(VbeMthDry(CurVbe))
End Sub

Sub Z_PjMthDry()
Brw DryFmtss(PjMthDry(CurPj))
End Sub


Function MdMthKy(A As CodeModule, Optional IsWrap As Boolean) As String()
Dim PjN$: PjN = MdPjNm(A)
Dim MdN$: MdN = MdNm(A)
MdMthKy = SrcMthKy(MdSrc(A), PjN, MdN, IsWrap)
End Function

Function ShfAs(A) As Variant()
Dim L$
L = LTrim(A)
If Left(L, 3) = "As " Then ShfAs = Array(True, LTrim(Mid(L, 4))): Exit Function
ShfAs = Array(False, A)
End Function

Function ShfTerm$(OLin$)
Dim L$, P%
L = LTrim(OLin)
If FstChr(L) = "[" Then
    P = SqBktEndPos(L)
    ShfTerm = Mid(L, 2, P - 2)
    OLin = LTrim(Mid(L, P + 1))
    Exit Function
End If
P = InStr(L, " ")
If P = 0 Then
    ShfTerm = L
    OLin = ""
    Exit Function
End If
ShfTerm = Left(L, P - 1)
OLin = Trim(Mid(L, P + 1))
End Function
Sub Z_VbeMthLinDry()
Brw DryFmtss(VbeMthLinDry(CurVbe))
End Sub
Function VbeMthLinDry(A As Vbe) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushAy VbeMthLinDry, PjMthLinDry(CvPj(P))
Next
End Function

Function PjMthLinDry(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushAy PjMthLinDry, MdMthLinDry(CvMd(M))
Next
End Function
Function MdMthLinDry(A As CodeModule) As Variant()
MdMthLinDry = SrcMthLinDry(MdBdyLy(A))
End Function

Sub PushNonBlankAy(O, M)
If Sz(M) > 0 Then Push O, M
End Sub

Function SplitComma(A) As String()
SplitComma = Split(A, ",")
End Function

Sub Z_VbeMthLinDryWP()
Brw DryFmtssWrp(VbeMthLinDryWP(CurVbe))
End Sub

Function VbeMthLinDryWP(A As Vbe) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushIAy VbeMthLinDryWP, PjMthLinDryWP(CvPj(P))
Next
End Function

Function PjMthLinDryWP(A As VBProject) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A))
    PushIAy PjMthLinDryWP, MdMthLinDryWP(CvMd(M))
Next
End Function

Function MdMthLinDryWP(A As CodeModule) As Variant()
MdMthLinDryWP = SrcMthLinDryWP(MdBdyLy(A))
End Function

Function LinMthDrWP(A) As Variant()
Dim Dr()
Dr = LinMthDr(A)
If Sz(Dr) = 0 Then Exit Function
Dr(3) = AyAddCommaSpcSfxExptLas(AyTrim(SplitComma(Dr(3))))
LinMthDrWP = Dr
End Function

Sub MthDrAsg(A, OShtMdy$, OShtTy$, ONm$, OPrm$, ORet$, OLinRmk$)
AyAsg A, OShtMdy, OShtTy, ONm, OPrm, ORet, OLinRmk
End Sub

Sub SrcMthDrAsg(A, OShtMdy$, OShtTy$, ONm$, OPrm$, ORet$, OLinRmk$, OLines$, OTopRmk$)
AyAsg A, OShtMdy, OShtTy, ONm, OPrm, ORet, OLinRmk, OLines, OTopRmk
End Sub

Function SrcMthFullDry(A$()) As Variant()
Dim Ix
For Each Ix In AyNz(SrcMthIx(A))
    PushI SrcMthFullDry, SrcMthIxFullDr(A, Ix)
Next
Dim Dr(): GoSub X
If Sz(Dr) > 0 Then
    PushI SrcMthFullDry, Dr
End If
Exit Function
X:
    Dim Dcl$, Cnt%
    Dcl = SrcDclLines(A)
    Cnt = LinCnt(Dcl)
    Const Fldss$ = "Ty Nm Cnt Lines"
    Dim Vy(): Vy = Array("Dcl", "*Dcl", Cnt, Dcl)
    If Dcl = "" Then
        Erase Dr
    Else
        Dr = VyDr(Vy, Fldss, SrcMthIxFullDrFny)
    End If
    Return
End Function

Function MdMthFullDrsFny() As String()
MdMthFullDrsFny = AyAdd(SslSy("PjFfn Pj MdTy Md"), SrcMthIxFullDrFny)
End Function

Function MdMthFullDrs(A As CodeModule, Optional B As WhMth) As Drs
Set MdMthFullDrs = Drs(MdMthFullDrsFny, MdMthFullDry(A, B))
End Function

Function PjMthFullDry(A As VBProject, Optional B As WhMdMth) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A, WhMdMth_Md(B)))
    PushIAy PjMthFullDry, MdMthFullDry(CvMd(M), WhMdMth_Mth(B))
Next
End Function

Function PjMthFullDrs(A As VBProject, Optional B As WhMdMth) As Drs
Dim O As Drs
Set O = Drs(MdMthFullDrsFny, PjMthFullDry(A, B))
Set O = DrsAddValIdCol(O, "Lines", "Pj")
Set O = DrsAddValIdCol(O, "Nm", "PjMth")
Set PjMthFullDrs = O
End Function

Function VbeMthFullDrs(A As Vbe, Optional B As WhPjMth) As Drs
Dim P, Fst As Boolean
Fst = True
For Each P In AyNz(VbePjAy(A, WhPjMth_Nm(B)))
    Dim M As Drs: Set M = PjMthFullDrs(CvPj(P), WhPjMth_MdMth(B))
    If Fst Then
        Set VbeMthFullDrs = M
        Fst = False
    Else
        PushDrs VbeMthFullDrs, M
    End If
Next
End Function

Function FbMthFullDrs(A, Optional B As WhPjMth) As Drs
If False Then
    Set FbMthFullDrs = VbeMthFullDrs(FbAcs(A).Vbe, B)
    Exit Function
End If
Dim Acs As New Access.Application
Debug.Print "FbMthFullDry: "; Now; " Start get Drs "; A; "==============="
Debug.Print "FbMthFullDry: "; Now; " Start open"
Set Acs = FbAcs(A)
Debug.Print "FbMthFullDry: "; Now; " Start get Drs "
Set FbMthFullDrs = VbeMthFullDrs(Acs.Vbe, B)
Debug.Print "FbMthFullDry: "; Now; " Start quit acs "
Acs.Quit acQuitSaveNone
Debug.Print "FbMthFullDry: "; Now; " acs is quit"
Set Acs = Nothing
Debug.Print "FbMthFullDry: "; Now; " acs is nothing"
End Function

Function FbvbeAyMthFullDrs(FbvbeAy(), Optional B As WhPjMth) As Drs
Dim I, Fst As Boolean
Fst = True
For Each I In FbvbeAy
    Dim A As Drs: GoSub X_A
    If Fst Then
        Set FbvbeAyMthFullDrs = A
        Fst = False
    Else
        PushDrs FbvbeAyMthFullDrs, A
    End If
Next
Exit Function
X_A:
    Select Case True
    Case IsStr(I):            Set A = FbMthFullDrs(I, B)
    Case TypeName(I) = "VBE": Set A = VbeMthFullDrs(CvVbe(I), B)
    Case Else: Stop
    End Select
    Return
End Function

Sub AAA()
Z_MthFullWbFmt
End Sub

Sub AAAA()
Z_UsrEdtMthLocDrs
End Sub

Function WbFstWs(A As Workbook) As Worksheet
Dim Ws
For Each Ws In A.Sheets
    Set WbFstWs = Ws
    Exit Function
Next
End Function

Sub Z_MthFullWbFmt()
Dim Wb As Workbook
Const Fx$ = "C:\Users\user\Desktop\Vba-Lib-1\Mth.xlsx"
MthFullWbFmt WbVis(FxWb(Fx))
Stop
End Sub
Function RgLasRow&(A As Range)
RgLasRow = A.Row + A.Rows.Count - 1
End Function
Function RgLasCol%(A As Range)
RgLasCol = A.Column + A.Columns.Count - 1
End Function
Function PtWs(A As PivotTable) As Worksheet
Set PtWs = A.Parent
End Function
Function PtCpyToLo(A As PivotTable, At As Range, Optional LoNm$) As ListObject
Dim R1, R2, C1, C2, NC, NR
    R1 = A.RowRange.Row
    C1 = A.RowRange.Column
    R2 = RgLasRow(A.DataBodyRange)
    C2 = RgLasCol(A.DataBodyRange)
    NC = C2 - C1 + 1
    NR = R2 - C1 + 1
WsRCRC(PtWs(A), R1, C1, R2, C2).Copy
At.PasteSpecial xlPasteValues

Set PtCpyToLo = RgLo(RgRCRC(At, 1, 1, NR, NC), LoNm)
End Function
Function WbCdNmWs(A As Workbook, CdNm$) As Worksheet
Dim Ws As Worksheet
For Each Ws In A.Sheets
    If Ws.CodeName = CdNm Then Set WbCdNmWs = Ws: Exit Function
Next
End Function
Function MthFullWbFmt(A As Workbook) As Workbook
Dim Ws As Worksheet, Lo As ListObject
Set Ws = WbCdNmWs(A, "MthLoc"): If IsNothing(Ws) Then Stop
Set Lo = WsLo(Ws, "T_MthLoc"): If IsNothing(Lo) Then Stop
Dim Ws1 As Worksheet:  GoSub X_Ws1
Dim Pt1 As PivotTable: GoSub X_Pt1
Dim Lo1 As ListObject: GoSub X_Lo1
Dim Pt2 As PivotTable: GoSub X_Pt2
Dim Lo2 As ListObject: GoSub X_Lo2
Ws1.Outline.ShowLevels , 1
Set MthFullWbFmt = WsWb(Ws)
Exit Function
X_Ws1:
    Set Ws1 = WbAddWs(WsWb(Ws))
    Ws1.Outline.SummaryColumn = xlSummaryOnLeft
    Ws1.Outline.SummaryRow = xlSummaryBelow
    Return
X_Pt1:
    Set Pt1 = LoPt(Lo, WsA1(Ws1), "MdTy Nm VbeLinesId Lines", "Pj")
    PtSetRowssOutLin Pt1, "Lines"
    PtSetRowssColWdt Pt1, "VbeLinesId", 12
    PtSetRowssColWdt Pt1, "Nm", 30
    PtSetRowssRepeatLbl Pt1, "MdTy Nm"
    Return
X_Lo1:
    Set Lo1 = PtCpyToLo(Pt1, Ws1.Range("G1"), LoNm:="T_MthLines")
    LoSetColWdt Lo1, "Nm", 30
    LoSetColWdt Lo1, "Lines", 100
    LoSetOutLin Lo1, "Lines"
    
    Return
X_Pt2:
    Set Pt2 = LoPt(Lo1, Ws1.Range("M1"), "MdTy Nm", "Lines")
    PtSetRowssRepeatLbl Pt2, "MdTy"
    Return
X_Lo2:
    Set Lo2 = PtCpyToLo(Pt2, Ws1.Range("Q1"), "T_UsrEdtMthLoc")
    Return
Set MthFullWbFmt = A
End Function
Sub Z_CurFbvbeAyMthFullWs()
WsVis CurFbvbeAyMthFullWs
End Sub

Sub Z_FbvbeAyMthFullWb()
Dim A()
    PushObj A, CurVbe
WbVis FbvbeAyMthFullWb(A, WhPjMth(MdMth:=WhMdMth(WhMd("Cls"))))
End Sub

Function PtPf(A As PivotTable, F) As PivotField
Set PtPf = A.PivotFields(F)
End Function

Function PtRowFldEntCol(A As PivotTable, F) As Range
Set PtRowFldEntCol = RgR(PtPf(A, F).DataRange, 1).EntireColumn
End Function

Sub PtSetRowssOutLin(A As PivotTable, Rowss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim F
For Each F In AyNz(SslSy(Rowss))
    PtRowFldEntCol(A, F).OutlineLevel = Lvl
Next
End Sub
Function LoEntCol(A As ListObject, C) As Range
Set LoEntCol = A.ListColumns(C).Range.EntireColumn
End Function
Sub LoSetColWdt(A As ListObject, Colss$, ColWdt As Byte)
If ColWdt <= 1 Then Stop
Dim C
For Each C In AyNz(SslSy(Colss))
    LoEntCol(A, C).ColumnWidth = ColWdt
Next
End Sub
Sub LoSetOutLin(A As ListObject, Colss$, Optional Lvl As Byte = 2)
If Lvl <= 1 Then Stop
Dim C
For Each C In AyNz(SslSy(Colss))
    LoEntCol(A, C).OutlineLevel = Lvl
Next
End Sub
Sub PtSetRowssColWdt(A As PivotTable, Rowss$, ColWdt As Byte)
If ColWdt <= 1 Then Stop
Dim F
For Each F In AyNz(SslSy(Rowss))
    PtRowFldEntCol(A, F).ColumnWidth = ColWdt
Next
End Sub
Sub PtSetRowssRepeatLbl(A As PivotTable, Rowss$)
Dim F
For Each F In AyNz(SslSy(Rowss))
    PtPf(A, F).RepeatLabels = True
Next
End Sub

Function FbvbeAyMthFullWb(FbvbeAy(), Optional B As WhPjMth) As Workbook
Set FbvbeAyMthFullWb = MthFullWbFmt(WsWb(FbvbeAyMthFullWs(FbvbeAy, B)))
End Function

Function CurFbvbeAyMthFullWs() As Worksheet
Set CurFbvbeAyMthFullWs = FbvbeAyMthFullWs(CurFbvbeAy)
End Function

Function FbvbeAyMthFullWs(FbvbeAy(), Optional B As WhPjMth) As Worksheet
Dim O As Drs
Set O = FbvbeAyMthFullDrs(FbvbeAy, B)
Set O = DrsAddValIdCol(O, "Nm", "VbeMth")
Set O = DrsAddValIdCol(O, "Lines", "Vbe")
Set FbvbeAyMthFullWs = WsSetCdNmAndLoNm(DrsWs(O), "MthLoc")
End Function

Private Sub Z_VyDr()
Dim Fny$(), Fldss$, Vy()
Fny = SslSy("A B C D E f")
Fldss = "C E"
Vy = Array(1, 2)
Ept = Array(Empty, Empty, 1, Empty, 2)
GoSub Tst
Exit Sub
Tst:
    Act = VyDr(Vy, Fldss, Fny)
    C
    Return
End Sub
Function VyDr(A(), Fldss$, Fny$()) As Variant()
Dim IxAy&(), U%
    IxAy = AyIxAy(Fny, SslSy(Fldss))
    U = AyMax(IxAy)
    GoSub X_ChkIxAy
Dim O(), J%, Ix
ReDim O(U)
For Each Ix In IxAy
    O(Ix) = A(J)
    J = J + 1
Next
VyDr = O
Exit Function
X_ChkIxAy:
    For Each Ix In IxAy
        If Ix <= -1 Then Stop
    Next
    Return
End Function

Sub Z_MdMthFullDrs()
DrsBrw MdMthFullDrs(CurMd)
End Sub

Function MdMthFullDry(A As CodeModule, Optional B As WhMth) As Variant()
Dim P As VBProject, Ffn$, Pj$, ShtTy$, Md$, MdTy$
Set P = MdPj(A)
Ffn$ = PjFfn(P)
Pj = P.Name
MdTy = MdShtTy(A)
Md = MdNm(A)
MdMthFullDry = DryInsC4(SrcMthFullDry(MdBdyLy(A)), Ffn, Pj, MdTy, Md)
End Function

Function MdShtTy$(A As CodeModule)
MdShtTy = CmpTyToSht(A.Parent.Type)
End Function

Function SrcMthIxFullDrFny() As String()
SrcMthIxFullDrFny = AyAdd(LinMthDrFny, SslSy("Lno Cnt Lines TopRmk"))
End Function
Sub Z_SrcMthNy()
Brw SrcMthNy(CurSrc)
End Sub
Function SrcMthNy(A$(), Optional B As WhMth) As String()
Dim L
For Each L In AyNz(A)
    PushNonBlankStr SrcMthNy, LinMthNm(L, B)
Next
End Function

Function SrcMthIxFullDr(A$(), MthIx) As Variant()
Dim L$, Lines$, TopRmk$, Lno%, Cnt%
    L = SrcContLin(A, MthIx)
    Lno = MthIx + 1
    Lines = SrcMthIxLines(A, MthIx)
    Cnt = SubStrCnt(Lines, vbCrLf) + 1
    TopRmk = SrcMthIxTopRmk(A, MthIx)
Dim Dr(): Dr = LinMthDr(L): If Sz(Dr) = 0 Then Stop
SrcMthIxFullDr = AyAdd(Dr, Array(Lno, Cnt, Lines, TopRmk))
End Function

Function LinMthDr(A) As Variant()
Dim L$, Mdy$, Ty$, Nm$, Prm$, Ret$, LinRmk$
L = A
Mdy = ShfShtMdy(L)
Ty = ShfMthTy(L): If Ty = "" Then Exit Function
Ty = MthShtTy(Ty)
Nm = ShfNm(L)
Ret = ShfMthSfx(L)
Prm = ShfBktStr(L)
If ShfX(L, "As") = "As" Then
    If Ret <> "" Then Stop
    Ret = ShfTerm(L)
End If
If ShfX(L, "'") = "'" Then
    LinRmk = L
End If
LinMthDr = Array(Mdy, Ty, Nm, Prm, Ret, LinRmk)
End Function
Function LinMthDrFny() As String()
LinMthDrFny = SslSy("Mdy Ty Nm Prm Ret LinRmk")
End Function
Function ShfRmk(A) As String()
Dim L$
L = LTrim(A)
If FstChr(L) = "'" Then
    ShfRmk = ApSy(Mid(L, 2), "")
Else
    ShfRmk = ApSy("", A)
End If
End Function

Function PjMthDry(A As VBProject, Optional B As WhMdMth, Optional C As MthBrkOpt) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A, WhMdMth_Md(B)))
    PushAy PjMthDry, MdMthDry(CvMd(M), WhMdMth_Mth(B), C)
Next
End Function

Function PjMthKyWs(A As VBProject) As Worksheet
Set PjMthKyWs = WsVis(SqWs(PjMthKySq(A)))
End Function

Function PjMthWs(A As CodeModule) As Worksheet
Set PjMthWs = WsVis(SqWs(PjMthSq(A)))
End Function

Function CvVbe(A) As Vbe
Set CvVbe = A
End Function

Sub AcsCls(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
End Sub

Sub AcsQuit(A As Access.Application)
On Error Resume Next
A.CloseCurrentDatabase
A.Quit
Set A = Nothing
End Sub

Function CvAcs(A) As Access.Application
Set CvAcs = A
End Function

Function CurFbvbeAy() As Variant()
PushObj CurFbvbeAy, CurVbe
PushIAy CurFbvbeAy, AppFbAy
End Function

Function DrsInsCol(A As Drs, ColNm$, C) As Drs
Set DrsInsCol = Drs(AyIns(A.Fny, ColNm), DryInsCol(A.Dry, C))
End Function

Sub PushDrs(O As Drs, A As Drs)
If Not IsEq(O.Fny, A.Fny) Then Stop
Set O = Drs(O.Fny, CvAy(AyAddAp(O.Dry, A.Dry)))
End Sub

Function VbeMthDrs(A As Vbe, Optional B As WhMth, Optional C As MthBrkOpt) As Drs
Dim O As Drs, O1 As Drs, O2 As Drs

Set O = Drs("Pj Md Mdy Ty Nm Lines", VbeMthDry(A))
Set O1 = DrsAddValIdCol(O, "Nm")
Set O2 = DrsAddValIdCol(O1, "Lines")
Set VbeMthDrs = O2
End Function

Function MthBrkOptFny(A As MthBrkOpt) As String()
MthBrkOptFny = SslSy("MthIx Cnt Lin Pj Md")
End Function

Function VbeMthDry(A As Vbe, Optional B As WhMth, Optional C As MthBrkOpt) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushAy VbeMthDry, PjMthDry(CvPj(P), B, C)
Next
End Function

Function CurVbeMthWb() As Workbook
Set CurVbeMthWb = VbeMthWb(CurVbe)
End Function


Function VbeMthWb(A As Vbe) As Workbook
Set VbeMthWb = WbVis(WbSavAs(WsWb(VbeMthWs(A)), VbeMthFx))
End Function

Function LoWs(A As ListObject) As Worksheet
Set LoWs = A.Parent
End Function

Function CurVbeMthWs() As Worksheet
Set CurVbeMthWs = VbeMthWs(CurVbe)
End Function

Function VbeMthWs(A As Vbe) As Worksheet
Set VbeMthWs = DrsWs(VbeMthDrs(A))
End Function

Private Sub ZZ_SrcMthDry()
Dim A(): A = SrcMthDry(CurSrc)
Stop
End Sub

Private Sub ZZ_VbeMthWb()
WbVis VbeMthWb(CurVbe)
End Sub

Private Sub ZZ_VbeMthWs()
WsVis VbeMthWs(CurVbe)
End Sub

Function VbeMthFx$()
VbeMthFx = FfnNxt(CurPjPth & "VbeMth.xlsx")
End Function

