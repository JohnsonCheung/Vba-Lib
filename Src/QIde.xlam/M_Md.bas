Attribute VB_Name = "M_Md"
Option Explicit
Private Type Mth
    Md As CodeModule
    Nm As String
End Type

Private Type DicPair
    A As Dictionary
    B As Dictionary
End Type
Private Type DCRslt ' DicCmpRslt
    Nm1 As String
    Nm2 As String
    AExcess As Dictionary
    BExcess As Dictionary
    ADif As Dictionary
    BDif As Dictionary
    Sam As Dictionary
End Type
Enum eReportSortingOption
    eDifOnly = 1
    eSamOnly = 2
    eBothDifAndSam = 3
End Enum



Function MdBdyLines$(A As CodeModule)
MdBdyLines = SrcBdyLines(MdSrc(A))
End Function

Function MdBdyLno%(A As CodeModule)
MdBdyLno = MdDclLinCnt(A) + 1
End Function

Function MdBdyLnoCnt(A As CodeModule) As LnoCnt
MdBdyLnoCnt = SrcBdyLnoCnt(MdSrc(A))
End Function

Function MdBdyLy(A As CodeModule) As String()
MdBdyLy = SrcBdyLy(MdSrc(A))
End Function

Function MdCanHasCd(A As CodeModule) As Boolean
Select Case MdTy(A)
Case _
    vbext_ComponentType.vbext_ct_StdModule, _
    vbext_ComponentType.vbext_ct_ClassModule, _
    vbext_ComponentType.vbext_ct_Document, _
    vbext_ComponentType.vbext_ct_MSForm
    MdCanHasCd = True
End Select
End Function

Function MdCmp(A As CodeModule) As VBComponent
Set MdCmp = A.Parent
End Function

Function MdCmpTy(A As CodeModule) As vbext_ComponentType
MdCmpTy = MdCmp(A).Type
End Function

Function MdContLin$(A As CodeModule, Lno&)
Dim J&, L&
L = Lno
Dim O$: O = A.Lines(L, 1)
While LasChr(O) = "_"
    L = L + 1
    O = RmvLasChr(O) & A.Lines(L, 1)
Wend
MdContLin = O
End Function

Function MdCurMthNm$(A As CodeModule)
Dim L&
   Dim R1&, R2&, C1&, C2&
   A.CodePane.GetSelection R1, C1, R2, C2
   L = R1
Dim K As vbext_ProcKind
MdCurMthNm = A.ProcOfLine(L, K)
End Function

Function MdDclLinCnt%(A As CodeModule)
MdDclLinCnt = SrcDclLinCnt(MdSrc(A))
End Function

Function MdDclLines$(A As CodeModule)
MdDclLines = JnCrLf(MdDclLy(A))
End Function

Function MdDclLy(A As CodeModule) As String()
MdDclLy = SrcDclLy(MdSrc(A))
End Function

Function MdDftMthNm$(Optional A As CodeModule, Optional MthNm$)
If MthNm = "" Then
   MdDftMthNm = MdCurMthNm(DftMd(A))
Else
   MdDftMthNm = A
End If
End Function

Function MdEnmBdyLy(A As CodeModule, EnmNm$) As String()
MdEnmBdyLy = DclEnmBdyLy(MdDclLy(A), EnmNm)
End Function

'Function MdMthDrs(Optional WithBdyLy As Boolean, _
'    Optional WithBdyLines As Boolean) As Drs
'Dim O As Drs
'    O = SrcMthDrs(MdSrc(A), WithBdyLy, WithBdyLines)
'MdMthDrs = DrsAddConstCol(O, "MdNm", MdNm(A))
'End Function
Function MdEnmItmCnt(A As CodeModule) As SrcItmCnt

End Function

Function MdEnmMbrCnt%(A As CodeModule, EnmNm$)
MdEnmMbrCnt = Sz(MdEnmMbrLy(A, EnmNm))
End Function

Function MdEnmMbrLy(A As CodeModule, EnmNm$) As String()
Dim Ly$(), O$(), J%
Ly = MdEnmBdyLy(A, EnmNm)
If AyIsEmp(Ly) Then Exit Function
Dim L
For Each L In Ly
   If Not SrcLin_IsRmk(L) Then
    Stop '
'       If Not Lin(L).IsEmp Then
'           Push O, Ly(J)
'       End If
   End If
Next
MdEnmMbrLy = O
End Function

Function MdEnmNy(A As CodeModule) As String()
MdEnmNy = DclEnmNy(MdDclLy(A))
End Function

Function MdEnsMth(A As CodeModule, MthNm$, NewMthLines$)
Dim OldMthLines$: OldMthLines = MthBdyLines(A, MthNm)
If OldMthLines = NewMthLines Then
    Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is same", MthNm, MdNm(A))
End If
MthRmv A, MthNm
MdAppLines A, NewMthLines
Debug.Print FmtQQ("MdEnsMth: Mth(?) in Md(?) is replaced <=========", MthNm, MdNm(A))
End Function

Function MdExp(A As CodeModule)
Dim F$: F = MdSrcFfn(A)
A.Parent.Export F
Debug.Print MdNm(A)
End Function

Function MdIsCls(A As CodeModule) As Boolean
MdIsCls = MdTy(A) = vbext_ct_ClassModule
End Function

Function MdIsExist(MdNm$, A As VBProject) As Boolean
On Error GoTo X
'MdIsExist = DftPj(A).VBComponents(MdNm).Name = MdNm
Exit Function
X:
End Function

Function MdLasLin$(A As CodeModule)
Dim N%: N = MdNLin(A)
If N = 0 Then Exit Function
MdLasLin = A.Lines(N, 1)
End Function

Function MdLasLno&(A As CodeModule)
MdLasLno = A.CountOfLines
End Function

Function MdLin$(A As CodeModule, Lno&)
If Lno <= 0 Then Exit Function
With A
    If Lno <= .CountOfLines Then MdLin = .Lines(Lno, 1)
End With
End Function

Function MdLines$(A As CodeModule)
With A
    If .CountOfLines = 0 Then Exit Function
    MdLines = .Lines(1, .CountOfLines)
End With
End Function

Function MdLinesByLnoCnt$(A As CodeModule, LnoCnt As LnoCnt)
With LnoCnt
    If .Cnt <= 0 Then Exit Function
    MdLinesByLnoCnt = A.Lines(.Lno, .Cnt)
End With
End Function

Function MdLy(A As CodeModule) As String()
MdLy = SplitCrLf(MdLines(A))
End Function

Function MdMthLinAy(A As CodeModule) As String()
MdMthLinAy = SrcMthLinAy(MdSrc(A))
End Function

Function MdMthNy(A As CodeModule, Optional MthNmPatn$ = ".") As String()
MdMthNy = AySrt(SrcMthNy(MdSrc(A), MthNmPatn))
End Function

Function MdNEnm%(A As CodeModule)
MdNEnm = DclNEnm(MdDclLy(A))
End Function

Function MdNLin%(A As CodeModule)
MdNLin = A.CountOfLines
End Function

Function MdNMth%(A As CodeModule)
MdNMth = SrcNMth(MdSrc(A))
End Function

Function MdNTy%(A As CodeModule)
MdNTy = SrcNTy(MdDclLy(A))
End Function

Function MdNm$(A As CodeModule)
MdNm = A.Parent.Name
End Function

Function MdOptCmpDbLno%(A As CodeModule)
Dim Ay$(): Ay = MdDclLy(A)
Dim J%
For J = 0 To UB(Ay)
    If HasPfx(Ay(J), "Option Compare Database") Then MdOptCmpDbLno = J + 1: Exit Function
Next
End Function

Function MdPatnLy(A As CodeModule, Patn$) As String()
Dim Ix&(): Ix = AyWhPatnIx(MdLy(A), Patn)
Dim O$(), I, Md As CodeModule
Dim N$: N = MdNm(A)
If AyIsEmp(Ix) Then Exit Function
For Each I In Ix
   Push O, FmtQQ("MdGoLno ""?"",??' ?", N, I + 1, vbTab, A.Lines(I + 1, 1))
Next
MdPatnLy = O
End Function

Function MdPj(A As CodeModule) As VBProject
Set MdPj = A.Parent.Collection.Parent
End Function

Function MdPjNm$(A As CodeModule)
End Function

Function MdPrvMthNy(A As CodeModule) As String()
MdPrvMthNy = SrcPrvMthNy(MdSrc(A))
End Function

Function MdResLy(A As CodeModule, ResNm$, Optional ResPfx$ = "ZZRes") As String()
Dim Z$()
    Z = MthBdyLy(A, ResPfx & ResNm)
    If AyIsEmp(Z) Then
        Er "MdResLy", "{MthNm} in {Md} is not found", ResPfx & ResNm, MdNm(A)
    End If
    Z = AyRmvFstEle(Z)
    Z = AyRmvLasEle(Z)
    Stop '
'    Z = SyRmvFstChr(Z)
MdResLy = Z
End Function

Function MdResStr$(A As CodeModule, ResNm$)
MdResStr = JnCrLf(MdResLy(A, ResNm))
End Function

Function MdSrc(A As CodeModule) As String()
MdSrc = MdLy(A)
End Function

Function MdSrcFfn$(A As CodeModule)
Stop '
'MdSrcFfn = Pjx(MdPj(A)).SrcPth & MdSrcFn(A)
End Function

Function MdSrcFn$(A As CodeModule)
MdSrcFn = MdCmp(A).Name & MdSrcExt(A)
End Function

Function MdTy(A As CodeModule) As vbext_ComponentType
MdTy = A.Parent.Type
End Function

Function MdTyLno$(A As CodeModule, TyNm$)
MdTyLno = -1
End Function

Function MdTyNm(A As CodeModule)
MdTyNm = CmpTy_Str(MdTy(A))
End Function

Function MdTyNy(A As CodeModule, Optional TyNmPatn$ = ".") As String()
MdTyNy = AySrt(DclTyNy(MdDclLy(A), TyNmPatn))
End Function

Function MdTyRRCC(A As CodeModule, TyNm$) As RRCC
Dim R&, C1&, C2&
R = MdTyLno(A, TyNm)
If R > 0 Then
    Stop '
'    With SubStrPos(A.Lines(R, 1), TyNm)
'        C1 = .FmPos
'        C2 = .ToPos
'    End With
End If
'MdTyRRCC = NewRRCC(R, R, C1, C2)
End Function

Function MdTyStr$(A As CodeModule)
MdTyStr = CmpTy_Str(MdTy(A))
End Function

Function MdWin(A As CodeModule) As VBIDE.Window
Set MdWin = A.CodePane.Window
End Function



Property Get Md(MdNm$) As CodeModule
Set Md = PjMd(CurPj, MdNm)
End Property

Sub MdAppDclLin(A As CodeModule, DclLines$)
A.InsertLines A.CountOfDeclarationLines + 1, DclLines
Debug.Print FmtQQ("MdAppDclLin: Module(?) a DclLin is inserted", MdNm(A))
End Sub

Sub MdAppLines(A As CodeModule, Lines$)
If Lines = "" Then Exit Sub
Dim Bef%
    Bef = A.CountOfLines
If A.CountOfLines = 0 Then
    A.AddFromString Lines
Else
    A.InsertLines A.CountOfLines + 1, Lines
End If
Dim Aft%
    Aft = A.CountOfLines
Dim Exp%
Stop '
'    Exp = Bef + Vb.Lines(Lines).LinCnt
'If Exp <> Aft Then Debug.Print FmtQQ("MdAppLines Er(LinCnt Added is not expected): Bef[?] LinCnt[?]: Exp(Bef+LinCnt)[?] <> Aft[?] AftBdyLinCnt[?]", Bef, Vb.Lines(Lines).LinCnt, Exp, Aft, Vb.Lines(MdBdyLines(A)).LinCnt)
End Sub

Sub MdAppLy(A As CodeModule, Ly$())
MdAppLines A, JnCrLf(Ly)
End Sub

Sub MdClr(A As CodeModule, Optional IsSilent As Boolean)
With A
    If .CountOfLines = 0 Then Exit Sub
    If Not IsSilent Then Debug.Print FmtQQ("MdClr: Md(?) of lines(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub

Sub MdClrBdy(A As CodeModule, Optional IsSilent As Boolean)
Stop
With A
    If .CountOfLines = 0 Then Exit Sub
    Dim N%, Lno%
        Lno = MdBdyLno(A)
        N = .CountOfLines - Lno + 1
    If N > 0 Then
        If Not IsSilent Then Debug.Print FmtQQ("MdClrBdy: Md(?) of lines(?) from Lno(?) is cleared", MdNm(A), N, Lno)
        .DeleteLines Lno, N
    End If
End With
End Sub

Sub MdCpy(A As CodeModule, ToMdNm$)
Dim Pj As VBProject
Set Pj = MdPj(A)
Stop '
'If Pjx(A).HasMdNm(ToMdNm) Then
'    Er "MdCpy", "{Pj} already contains {ToMdNm}.  {Md} cannot be copied", MdPjNm(A), ToMdNm, MdNm(A)
'End If
Dim Ty As vbext_ComponentType: Ty = MdTy(A)
Dim O As CodeModule
'Set O = PjCrtMd(Pj, ToMdNm, Ty)
MdAppLines O, MdLines(A)
End Sub

Sub MdGo(A As CodeModule)
MdShw A
WinOf_BrwObj.Visible = True
WinAp_Keep MdWin(A), WinOf_BrwObj
WinOf_Imm_Cls
TileV
End Sub

Sub MdGoLno(A As CodeModule, Lno&)
Stop '
'MdGoRRCC A, NewRRCC(Lno, Lno, 1, 1)
End Sub

Sub MdGoRRCC(A As CodeModule, RRCC As RRCC)
Stop '
'If RRCC_IsEmp(RRCC) Then Debug.Print FmtQQ("Given RRCC_ is empty"): Exit Sub
MdShw A
If IsMdRRCCOutSideMd(RRCC, A) Then
    With RRCC
    Stop '
'        Debug.Print FmtQQ("MdGoRg: Given ? is outside given Md[?]-(MaxR ?)(MaxR1C ?)(MaxR2C ?)", RRCC_Str(RRCC), MdNm(A), MdNLin(A), Len(A.Lines(.R1, 1)), Len(A.Lines(.R2, 1)))
    End With
    Exit Sub
End If
With RRCC
    A.CodePane.SetSelection .R1, .C1, .R2, .C2
End With
End Sub

Sub MdGoTy(A As CodeModule, TyNm$)
MdGoRRCC A, MdTyRRCC(A, TyNm)
End Sub

Sub MdRen(A As CodeModule, NewNm$)
Const CSub$ = "MdRen"
Dim Nm$: Nm = MdNm(A)
If NewNm = Nm Then
    Debug.Print FmtQQ("MdRen: Given Md-[?] name and NewNm-[?] is same", Nm, NewNm)
    Exit Sub
End If
If MdIsExist(NewNm, MdPj(A)) Then
    Debug.Print FmtQQ("MdRen: Md-[?] already exist.  Cannot rename from [?]", NewNm, MdNm(A))
    Exit Sub
End If
MdCmp(A).Name = NewNm
Debug.Print FmtQQ("MdRen: Md-[?] renamed to [?] <==========================", Nm, NewNm)
End Sub

Sub MdRmv(A As CodeModule)
Dim C As VBComponent: Set C = A.Parent
C.Collection.Remove C
End Sub

Sub MdRmvBdy(A As CodeModule)
MdRmvLnoCnt A, MdBdyLnoCnt(A)
End Sub

Sub MdRmvEndBlankLines(A As CodeModule)
If A.CountOfLines = 0 Then Exit Sub
Dim J&
Dim HasRmv As Boolean
HasRmv = True
While HasRmv
    HasRmv = False
    J = J + 1
    If J > 100000 Then
        Stop
    End If
    If Trim(A.Lines(A.CountOfLines, 1)) = "" Then
        A.DeleteLines A.CountOfLines, 1
        HasRmv = True
    End If
Wend
End Sub

Sub MdRmvLines(A As CodeModule)
If A.CountOfLines = 0 Then Exit Sub
A.DeleteLines 1, A.CountOfLines
End Sub

Sub MdRmvLnoCnt(A As CodeModule, LnoCnt As LnoCnt)
With LnoCnt
    If .Cnt = 0 Then Exit Sub
    A.DeleteLines .Lno, .Cnt
End With
End Sub

Sub MdRmvLnoCntAy(A As CodeModule, LnoCntAy() As LnoCnt)
If Sz(LnoCntAy) = 0 Then Exit Sub
Dim J%, M&
M = LnoCntAy(0).Lno
For J = 1 To UB(LnoCntAy)
    If M > LnoCntAy(J).Lno Then Stop
    M = LnoCntAy(J).Lno
Next

For J = UB(LnoCntAy) To 0 Step -1
    MdRmvLnoCnt A, LnoCntAy(J)
Next
End Sub

Sub MdRmvNmPfx(A As CodeModule, Pfx$)
Dim Nm$: Nm = MdNm(A): If Not HasPfx(Nm, Pfx) Then Exit Sub
MdRen A, RmvPfx(MdNm(A), Pfx)
End Sub

Sub MdRmvOptCmpDb(A As CodeModule)
Dim I%: I = MdOptCmpDbLno(A)
If I = 0 Then Exit Sub
A.DeleteLines I
Debug.Print "MdRmvOptCmpDb: Option Compare Database at line " & I & " is removed"
End Sub

Sub MdRpl(A As CodeModule, NewMdLines$)
MdClr A
MdAppLines A, NewMdLines
End Sub

Sub MdRplBdy(A As CodeModule, NewMdBdy$)
MdClrBdy A
MdAppLines A, NewMdBdy
End Sub

Sub MdRplLin(A As CodeModule, Lno%, NewLin$)
With A
    .DeleteLines Lno
    .InsertLines Lno, NewLin
End With
End Sub

Sub MdShw(A As CodeModule)
A.CodePane.Show
End Sub

Sub MdSrtRptBrw(A As CodeModule)
Dim Old$: Old = MdBdyLines(A)
Dim NewLines$: NewLines = MdSrtedLines(A)
Dim O$: O = IIf(Old = NewLines, "(Same)", "<====Diff")
Debug.Print MdNm(A), O
End Sub




Private Sub MdAppLines__Tst()
Const MdNm$ = "Module1"
MdAppLines CurMd, "'aa"
End Sub

Sub MdDftMthNm__Tst()
Dim I, Md As CodeModule
For Each I In PjMbrAy(CurPj)
   MdShw CvMd(I)
   Debug.Print MdNm(Md), MdDftMthNm(Md)
Next
End Sub

Private Sub MdEnmMbrCnt__Tst()
'Ass MdEnmMbrCnt(Md("Ide"), "AA") = 1
End Sub

Private Sub MdLy__Tst()
aybrw MdLy(CurMd)
End Sub

Sub MdOptCmpDbLin__Tst()
Dim I, Md As CodeModule
For Each I In PjMbrAy(CurPj)
    Set Md = I
    Debug.Print MdNm(Md), MdOptCmpDbLno(Md)
Next
End Sub

Sub MdRen__Tst()
Stop '
'MdRen Md("A_Rs1"), "A_Rs"
End Sub

Sub MdRmvLnoCntAy__Tst()
Dim A() As LnoCnt
Stop
'A = MthLnoCntAy(Md("Md_"), "XXX")
'MdRmvLnoCntAy Md("Md_"), A
End Sub


Private Function MdSrcExt$(A As CodeModule)
Dim O$
Select Case MdCmpTy(A)
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function
Function MdMthDic(A As CodeModule) As Dictionary
Set MdMthDic = SrcDic(MdSrc(A))
End Function
Function MdMdLisDr(A As CodeModule) As Variant()
Dim O(), N$
Stop '
'N = MdNm(A)
'Push O, Pjx(MdPj(A)).Nm
'Push O, TakBef(N, "_")
'Push O, N
'Push O, MdTyNm(A)
'Push O, A.CountOfLines
'With MdMthCnt(A)
'   Push O, .N
'   Push O, .NPub
'   Push O, .NPrv
'End With
'With MdTyItmCnt(A)
'   Push O, .N
'   Push O, .NPub
'   Push O, .NPrv
'End With
'With MdEnmItmCnt(A)
'   Push O, .N
'   Push O, .NPub
'   Push O, .NPrv
'End With
'MdMdLisDr = O
End Function
Function MdMdLisDt(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt) As Dt
Dim Dt As Dt: ' Dt = MdMdInfDt(A, MdNmPatn)
Stop '
Select Case Srt
'Case elmsLines: Dt = DtSrt(Dt, "Lines", True)
'Case elmsMd: Dt = DtSrt(Dt, "Md")
'Case elmsNMth: Dt = DtSrt(Dt, "NMth", True)
Case Else: Stop
End Select
MdMdLisDt = Dt
End Function
Function MdMthCnt(A As CodeModule) As SrcItmCnt
Dim B As Drs: B = MdMthDrs(A)
Dim N%, NPub%, NPrv%
N = Sz(B.Dry)
Stop '
'NPub = DrsRowCnt(B, "Mdy", "Public") + DrsRowCnt(B, "Mdy", "")
'NPrv = DrsRowCnt(B, "Mdy", "Private")
MdMthCnt = NewSrcItmCnt(N, NPub, NPrv)
End Function
Function MdMthDrs(A As CodeModule, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
Set MdMthDrs = SrcMthDrs(MdSrc(A), MdNm(A), MdTy(A), WithBdyLy, WithBdyLines)
End Function
Function MdMthDry(A As CodeModule, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Stop '
'MdMthDry = SrcMthDry(MdSrc(A), MdNm(A), MdTyStr(A), WithBdyLy, WithBdyLines)
End Function
Function MdTyItmCnt(A As CodeModule) As SrcItmCnt

End Function
Sub MdMdLisDtBrw(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt)
Stop '
'DtBrw MdMdLisDt(A, MdNmPatn, Srt)
End Sub
Sub MdMdLisDtDmp(A As CodeModule, Optional MdNmPatn$ = ".", Optional Srt As eLisMdSrt)
Stop '
'DtDmp MdMdLisDt(A, MdNmPatn, Srt)
End Sub
