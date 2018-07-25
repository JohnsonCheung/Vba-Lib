Attribute VB_Name = "F__Tool"
Option Explicit
Sub LisA()
Lis_Mth "^AA"
End Sub

Function Shw_Pj_SrtRptWb(Optional PjNm$) As Workbook
PjSrtRptWb DftPj(PjNm), Vis:=True
End Function

Sub Add_Cls(Nm$)
PjAddMbr CurPj, Nm, vbext_ct_ClassModule
Shw_Mbr CurPjNm & "." & Nm
End Sub

Sub Add_Fun(FunNm$)
MdAddFun CurMd, FunNm, IsFun:=True
End Sub

Sub Add_Md(Nm$)
PjAddMbr CurPj, Nm, vbext_ct_StdModule
End Sub

Sub Add_Sub(SubNm$)
MdAddFun CurMd, SubNm, IsFun:=False
End Sub

Sub Add_VbeRf_QTool()
Dim I, P As VBProject
For Each I In VbePjAy(CurVbe)
    Set P = I
    PjAddRf P, "QTool"
Next
End Sub

Sub Brw_DupMdNm()
AyBrw VbeDupMdNy(CurVbe)
End Sub

Sub Brw_InproperMth()
Brw_Pj_InproperMth
End Sub

Sub Brw_Md_InproperMth()
AyBrw MdMthNyOfInproper(CurMd)
End Sub

Sub Brw_Md_Mth()
DicBrw MdDicOfMthKeyzzzMthLines(CurMd)
End Sub

Sub Brw_Md_MthKy()
AyBrw MdMthKy(CurMd, IsWrap:=True)
End Sub

Sub Brw_Md_MthNm(Optional MthNmPatn$, Optional Mdy0$)
AyBrw MdMthNy(CurMd, MthNmPatn, IsNoMdNmPfx:=True, Mdy0:=Mdy0)
End Sub

Sub Brw_Md_SrtRpt(Optional MdDNm0$)
Dim N$: N = DftMdDNm(MdDNm0)
AyBrw MdSrtRptLy(Md(N))
End Sub

Sub Brw_Pj_FunFNy()
AyBrw PjFunFNy(CurPj)
End Sub

Sub Brw_Pj_InproperMth()
AyBrw PjMthNyOfInproper(CurPj)
End Sub

Sub Brw_Pj_Mth(Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".")
AyBrw PjMthNy(CurPj, MthNmPatn:=MthNmPatn, MbrNmPatn:=MbrNmPatn)
End Sub

Sub Brw_Pj_MthKy()
AyBrw PjMthKy(CurPj, IsWrap:=True)
End Sub

Sub Brw_Pj_SrtRpt()
AyBrw PjSrtRptLy(CurPj)
End Sub

Sub Brw_Vbe_DupFunDrs(Optional IsSamMthBdyOnly As Boolean)
WsVis DrsWs(VbeDupFunDrs(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly))
End Sub

Sub Brw_Vbe_DupFunFNy(Optional IsSamMthBdyOnly As Boolean)
AyBrw VbeDupFunFNy(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Brw_Vbe_FunFNy()
AyBrw VbeFunFNy(CurVbe)
End Sub

Sub Brw_Vbe_InproperMth()
AyBrw VbeMthNyOfInproper(CurVbe)
End Sub

Sub Brw_Vbe_MthKy()
AyBrw VbeMthKy(CurVbe, IsWrap:=True)
End Sub

Sub Brw_Vbe_SrcPth()
VbeSrcPthBrw CurVbe
End Sub

Sub Brw_Vbe_SrtRpt()
AyBrw VbeSrtRptLy(CurVbe)
End Sub

Sub Cls_Win()
VbeClsWin CurVbe
End Sub

Sub Cls_Win_ExcptImm(Optional ExcptWinTyAy)
VbeClsWin CurVbe, Array(VBIDE.vbext_wt_Immediate)
End Sub

Sub Cmp_DupFun()
FunNm_Cmp CurMthNm
End Sub

Sub Cmp_Fun(Optional FunNm0$, Optional InclSam As Boolean)
FunNm_Cmp DftMthNm(FunNm0), InclSam
End Sub

Sub Cmp_Vbe_DupFun(Optional InclSam As Boolean)
AyBrw VbeDupFunCmpLy(CurVbe, InclSam:=InclSam)
End Sub

Sub Compile_Pj()
PjCompile CurPj
End Sub

Sub Compile_Vbe()
AyDo VbePjAy(CurVbe), "PjCompile"
End Sub

Sub Cpy_Mbr(FmPjMbrDotNm$)
MdCpy Md(FmPjMbrDotNm), CurPj
End Sub

Sub Cpy_Md_ToPj(ToPjNm$)
MdCpy CurMd, Pj(ToPjNm)
End Sub

Sub Dlt_Md()
If MsgBox(FmtQQ("Delete this Md[?]", CurMdNm), vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
CurPj.VBComponents.Remove CurCmp
End Sub

Sub Dmp_Vbe_DupFun(Optional InclSam As Boolean)
Dim Ay$(): Ay = VbeDupFunCmpLy(CurVbe, InclSam:=InclSam)
Dim Ay1$(): Ay1 = AyFstNEle(Ay, 100)
AyDmp Ay1
If Sz(Ay) > 100 Then
    Debug.Print "...." & Sz(Ay) - 100 & " more"
End If
End Sub

Sub Dmp_Vbe_DupFunFNy(Optional IsSamMthBdyOnly As Boolean)
AyDmp VbeDupFunFNy(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Export_Pj()
PjExport CurPj
End Sub

Sub Export_Vbe()
VbeExport CurVbe
End Sub

Sub Gen_Md_TstSub()
Md_Gen_TstSub CurMd
End Sub

Sub Gen_Pj_ConstructorMd()
Stop '
End Sub

Sub Gen_Pj_TstClass()
Pj_Gen_TstClass CurPj
End Sub

Sub Gen_Pj_TstSub()
Pj_Gen_TstSub CurPj
End Sub

Sub Gen_Vbe_TstClass()
End Sub

Sub Lis_CurMth()
Debug.Print MthLines(CurMth)
End Sub

Sub Lis_Md()
Dim A$()
    A = PjMbrNy(CurPj)
    A = AySrt(A)
    A = AyAddPfx(A, "Shw_Mbr """)
AyDmp A
End Sub

Sub Lis_Md_FunPfx()
AyDmp AySrt(MdFunPfxAy(CurMd))
End Sub

Sub Lis_Md_InproperMth(Optional MdDNm0$)
AyDmp MdMthNyOfInproper(Md(DftMdDNm(MdDNm0)))
End Sub

Sub Lis_Md_Mth(Optional MthNmPatn$ = ".", Optional Mdy0$)
AyDmp AyAddPfx(MdMthNy(CurMd, MthNmPatn, Mdy0:=Mdy0), CurPjNm & ".")
End Sub

Sub Lis_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy0$)
Lis_Vbe_Mth MthNmPatn, MdNmPatn, Mdy0
End Sub

Sub Lis_Pj()
Dim A$()
    A = VbePjNy(CurVbe)
    A = AyAddPfx(A, "Shw_Pj """)
AyDmp A
End Sub

Sub Lis_Pj_DupFunFNy(Optional IsSamMthBdyOnly As Boolean)
AyDmp PjDupFunFNy(CurPj, IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Lis_Pj_FunPfx()
AyDmp AySrt(PjFunPfxAy(CurPj))
End Sub

Sub Lis_Vbe_FunPfx()
AyDmp AySrt(VbeFunPfxAy(CurVbe))
End Sub

Sub Lis_Pj_InproperMth(Optional PjNm0$)
If PjNm0 <> "" Then Shw_Pj PjNm0
AyDmp AyAddPfxSfx(PjMthNyOfInproper(CurPj), "MthMovToProperMd MthDNm_Mth(""", """)")
End Sub

Sub Lis_Pj_Mth(Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".", Optional Mdy0$)
Dim A$()
    A = PjMthNy(CurPj, MthNmPatn:=MthNmPatn, MbrNmPatn:=MbrNmPatn, Mdy0:=Mdy0)
    A = AySrt(A)
    A = AyAddPfx(A, "Shw_Mth """)
AyDmp A
End Sub

Sub Lis_Vbe_DupFunFNy(Optional IsSamMthBdyOnly As Boolean)
Dim A$(): A = VbeDupFunFNy(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
Dim A1$(): A1 = AyDblQuote(A)
AyDmp AyAddPfx(A1, "Shw ")
End Sub

Sub Lis_Vbe_InproperMth()
AyDmp VbeMthNyOfInproper(CurVbe)
End Sub

Sub Lis_Vbe_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$)
Dim A$()
    A = VbeMthNy(CurVbe, MthNmPatn, MdNmPatn, Mdy)
    A = AySrt(A)
AyDmp AyAddPfx(A, "Shw """)
End Sub

Sub Mov_Fun(Optional MthDNm0$)
'Mov Fun to its proper-module
'Fun here means Public-Prp/Sub/Fun, in a Md, not class
'         or    Private-Sub ZZ_xxx, in a Md, not class
'proper-module means, M_Xxx where Xxx is function-MdPfx
'MdPfx-of-a-fun is a Pfx of a funNm which is used to give a proper-module-nm of M_Xxx
MthMovToProperMd DftMth(MthDNm0)
End Sub

Sub Mov_Fun_ToProperMd()
'Move all Inproper-Fun in CurMd to its proper module in same Pj
'If non-exist-inproper-module will be created
'If a Fun in a module of name of format M_XXX,
'   if the Fun-name-pfx is not XXX, => it is inproper-fun
'else
'   => it is proper-fun
Dim I, M As CodeModule, Ny$()
Set M = CurMd
Ny = MdMthNyOfInproper(CurMd)
If Sz(Ny) = 0 Then Exit Sub
Dim N
Dim Mth As New Mth
Set Mth.Md = M
For Each N In Ny
    Mth.Nm = N
    MthMovToProperMd Mth
Next
End Sub

Sub Mov_MbrPatn_ToPj(MbrNmPatn$, ToPjNm$)
Dim Ay() As CodeModule: Ay = PjMbrAy(CurPj, MbrNmPatn)
If Sz(Ay) = 0 Then Exit Sub
Dim I, P As VBProject
Set P = Pj(ToPjNm)
For Each I In Ay
    Md_Mov_ToPj CvMd(I), P
Next
Cls_Win
End Sub

Sub Mov_Md_ToPj(ToPjNm$)
If CurPjNm = ToPjNm Then
    Debug.Print FmtQQ("Mov_Md: ToPjNm(?) cannot be CurPjNm", ToPjNm)
    Exit Sub
End If
Md_Mov_ToPj CurMd, Pj(ToPjNm)
End Sub

Sub Ren_Md(NewNm$)
If PjHasCmp(CurPj, NewNm) Then
    MsgBox FmtQQ("Md(?) exists in CurPj(?).  Cannot rename.", NewNm, CurPjNm), , "M_A:RenMd"
    Exit Sub
End If
CurMd.Name = NewNm
End Sub

Sub Rmk_All()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In PjMbrAy(CurPj)
    Set Md = I
    If MdRmk(Md) Then
        NRmk = NRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub

Sub Rmk_Mth()
Dim W As VBIDE.Window
Set W = CurCdWin
MthRmk CurMth
WinOf_Imm.Close
W.SetFocus
End Sub

Sub Sav_Pj()
PjSav CurPj
End Sub

Sub Sav_Vbe()
Dim V As Vbe: Set V = CurVbe
VbeSav V
DoEvents
VbeDmpIsSaved V
End Sub

Sub Shw(XNm$)
If IsMthFNm(XNm) Then
    Shw MthFNm_MthDNm(XNm)
End If
Dim A$(): A = Split(XNm, ".")
Select Case Sz(A)
Case 1
    Select Case True
    Case Left(XNm, 1) = "Q":  Shw_Pj XNm
    Case Else
        If IsMdNm(XNm) Then
            Shw_Mbr XNm
        Else
            Shw_Mth XNm
        End If
    End Select
Case 2
    Select Case True
    Case Left(A(0), 1) = "Q"
        If IsMdNm(A(1)) Then
            Shw_Mbr XNm
        Else
            Shw_Mth XNm
        End If
    Case IsMdNm(A(0))
        Shw_Mth XNm
    Case Else
        Debug.Print "For 2 Segment, 1st Segment of {Q* M_* S_* F_* G_*}"
        Stop
    End Select
Case 3
    Shw_Mth XNm
Case Else
Debug.Print "XNm has " & Sz(A) & " segments.  It should be 1 2 or 3"
End Select
End Sub

Sub Shw_Mbr(MdXNm$)
Dim E As Either
E = MdXNm_Either(MdXNm)
If E.IsLeft Then
    MdGo Md(E.Left)
    Exit Sub
End If
Dim Ny$()
    Ny = E.Right
If Sz(Ny) = 0 Then
    Debug.Print MdXNm; "<-- No such module"
    Exit Sub
End If
Dim I
For Each I In Ny
    Debug.Print "Shw_Mbr """; I; "."; MdXNm
Next
End Sub

Sub Shw_Mth(Mth_DNm_or_FNm$)
Dim D$
If IsMthFNm(Mth_DNm_or_FNm) Then
    D = MthFNm_MthDNm(Mth_DNm_or_FNm)
Else
    D = Mth_DNm_or_FNm
End If
Dim M As Mth
Set M = MthDNm_Mth(D)
MdGoLCCOpt M.Md, MthLCCOpt(M)
End Sub
Sub A1()

End Sub
Sub Shw_Pj(PjNm$)
PjGo Pj(PjNm)
End Sub

Sub Srt_G_Tool()
Dim M As CodeModule: Set M = Md("QTool.G_Tool")
Dim Src$(): Src = MdSrc(M)
Dim Cxt$: Cxt = SrcSrtedLines(Src)
If Cxt = Join(Src, vbCrLf) Then
    Debug.Print "Md(F__Tool) is alread sorted"
Else
    MdRplCxt M, Cxt
End If
End Sub

Sub Srt_Md(Optional MdNm$)
MdSrt DftMd(MdNm)
End Sub

Sub Srt_Pj()
PjSrt CurPj
End Sub

Sub Srt_Vbe()
VbeSrt CurVbe
End Sub

Sub Sync_Fun(Optional FunFNm0$)
Dim M As Mth
If FunFNm0 = "" Then
    Set M = CurMth
Else
    If Not IsMthFNm(FunFNm0) Then Stop
    Set M = MthFNm_Mth(FunFNm0)
End If
FunSync M, ShwCmpLyAft:=True
End Sub

Sub UnRmk_All()
Dim I, Md As CodeModule
Dim NUnRmk%, Skip%
For Each I In PjMbrAy(CurPj)
    Set Md = I
    If MdUnRmk(Md) Then
        NUnRmk = NUnRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NUnRmk"; NUnRmk
Debug.Print "SKip"; Skip
End Sub

Sub UnRmk_Mth()
MthUnRmk CurMth
WinOf_Imm.Close
End Sub

Sub Wb_Vbe_Mth(Optional InclMthLines As Boolean)
VbeMthWb CurVbe, InclMthLines
End Sub

Sub Ws_Md_MthKy()
MdMthWs CurMd
End Sub

Sub Ws_Pj_Mth()
PjMthWs CurPj
End Sub

Sub Ws_Vbe_Mth(Optional InclMthLines As Boolean)
VbeMthWs CurVbe, InclMthLines
End Sub

Sub Srt_F__Tool()
Dim M As CodeModule: Set M = Md("QTool.F__Tool")
Dim Src$(): Src = MdSrc(M)
Dim Cxt$: Cxt = SrcSrtedLines(Src)
If Cxt = Join(Src, vbCrLf) Then
    Debug.Print "Md(F__Tool) is already sorted"
Else
    MdRplCxt M, Cxt
End If
End Sub
