Attribute VB_Name = "F_Tool"
Option Explicit

Sub Add_Cls(Nm$)
PjAddMbr CurPj, Nm, vbext_ct_ClassModule
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
For Each I In CurVbe_PjAy
    Set P = I
    PjAddRf P, "QTool"
Next
End Sub

Sub Brw_DupMdNm()
AyBrw CurVbe_DupMdNy
End Sub

Sub Brw_InproperMth()
Brw_Pj_InproperMth
End Sub

Sub Brw_Md_InproperMth()
AyBrw MdMthNy_OfInproper(CurMd)
End Sub

Sub Brw_Md_Mth()
DicBrw Md_Dic_Of_MthKey_MthLines(CurMd)
End Sub
Property Get Md_Dic_Of_MthKey_MthLines(A As CodeModule) As Dictionary
Set Md_Dic_Of_MthKey_MthLines = Src_Dic_Of_MthKey_MthLines(MdSrc(A), MdPjNm(A), MdNm(A))
End Property
Sub Brw_Md_MthKy()
AyBrw MdMthKy(CurMd, IsSngLinFmt:=True)
End Sub

Sub Brw_Md_MthNm(Optional MthNmPatn$, Optional Mdy0$)
AyBrw MdMthNy(CurMd, MthNmPatn, IsNoMdNmPfx:=True, Mdy0:=Mdy0)
End Sub

Sub Brw_Pj_MthFNy()
AyBrw PjMthFNy(CurPj)
End Sub

Sub Brw_Pj_InproperMth()
AyBrw PjMthNy_OfInproper(CurPj)
End Sub

Sub Brw_Pj_Mth(Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".")
AyBrw PjMthNy(CurPj, MthNmPatn:=MthNmPatn, MbrNmPatn:=MbrNmPatn)
End Sub

Sub Brw_Pj_MthKy()
AyBrw PjMthKy(CurPj, IsSngLinFmt:=True)
End Sub

Sub Brw_Pj_SrtRpt()
AyBrw PjSrtRptLy(CurPj)
End Sub

Sub Brw_Md_SrtRpt(Optional MdDNm0$)
Dim N$: N = DftMdDNm(MdDNm0)
AyBrw MdSrtRptLy(Md(N))
End Sub

Sub Brw_Vbe_DupMthDrs(Optional IsSamMthBdyOnly As Boolean)
WsVis DrsWs(VbeDupMthDrs(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly))
End Sub

Sub Brw_Vbe_DupMthFNy(Optional IsSamMthBdyOnly As Boolean)
AyBrw VbeDupMthFNy(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Brw_Vbe_DupFun()
AyBrw VbeDupFunLy(CurVbe)
End Sub

Sub Brw_Vbe_MthFNy()
AyBrw VbeMthFNy(CurVbe)
End Sub

Sub Brw_Vbe_InproperMth()
AyBrw VbeMthNy_OfInproper(CurVbe)
End Sub

Sub Brw_Vbe_MthKy()
AyBrw VbeMthKy(CurVbe, IsSngLinFmt:=True)
End Sub

Sub Brw_Vbe_SrcPth()
VbeSrcPthBrw CurVbe
End Sub

Sub Brw_Vbe_SrtRpt()
AyBrw VbeSrtRptLy(CurVbe)
End Sub

Sub Cls_Win()
Dim W As VBIDE.Window
For Each W In CurVbe.Windows
    W.Close
Next
End Sub

Sub Compile_Pj()
PjCompile CurPj
End Sub

Sub Compile_Vbe()
AyDo CurVbe_PjAy, "PjCompile"
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

Sub Dmp_CurMth()
Debug.Print MthLines(CurMth)
End Sub

Sub Dmp_Md_InproperMth()
AyDmp MdMthNy_OfInproper(CurMd)
End Sub

Sub Dmp_Pj_DupMthFNy(Optional IsSamMthBdyOnly As Boolean)
AyDmp PjDupMthFNy(CurPj, IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub
Sub Dmp_Pj_InproperMth(Optional PjNm0$)
If PjNm0 <> "" Then Shw_Pj PjNm0
AyDmp PjMthNy_OfInproper(CurPj)
End Sub

Sub Dmp_Vbe_DupMthFNy(Optional IsSamMthBdyOnly As Boolean)
AyDmp VbeDupMthFNy(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Dmp_Vbe_InproperMth()
AyDmp VbeMthNy_OfInproper(CurVbe)
End Sub

Sub Export_Pj()
PjExport CurPj
End Sub

Sub Export_Vbe()
VbeExport CurVbe
End Sub

Sub Gen_Md_TstSub()
MdGen_TstSub CurMd
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

Sub Shw_Mbr(PjMbrDotNm$)
Dim E As Either
E = PjMbrDotNm_Either(PjMbrDotNm)
If E.IsLeft Then
    MdGo Md(E.Left)
    Exit Sub
End If
Dim Ny$()
    Ny = E.Right
If Sz(Ny) = 0 Then
    Debug.Print PjMbrDotNm; "<-- No such module"
    Stop '
    Exit Sub
End If
Dim I
For Each I In Ny
    Debug.Print "Shw_Mbr """; I; "."; PjMbrDotNm
Next
End Sub
Sub Shw(Nm$)
Dim A$(): A = Split(Nm, ".")
Select Case Sz(A)
Case 1
    Select Case True
    Case Left(Nm, 1) = "Q":  Shw_Pj Nm
    Case Else
        If IsMdNm(Nm) Then
            Shw_Mbr Nm
        Else
            Shw_Mth Nm
        End If
    End Select
Case 2
    Select Case True
    Case Left(A(0), 1) = "Q"
        If IsMdNm(A(1)) Then
            Shw_Mbr Nm
        Else
            Shw_Mth Nm
        End If
    Case IsMdNm(A(0))
        Shw_Mth Nm
    Case Else
        Debug.Print "For 2 Segment, 1st Segment of {Q* M_* S_* F_* G_*}"
        Stop
    End Select
Case 3
    Shw_Mth Nm
Case Else
Debug.Print "Nm has " & Sz(A) & " segments"
End Select
End Sub
Sub Shw_Mth(MthDNm$)
Dim M As Mth
Set M = MthDNm_Mth(MthDNm)
MdGoLCCOpt M.Md, MthLCCOpt(M)
End Sub

Sub Shw_Pj(PjNm$)
PjGo Pj(PjNm)
End Sub

Sub Lis_Md()
Dim A$()
    A = PjMbrNy(CurPj)
    A = AySrt(A)
    A = AyAddPfx(A, "Shw_Mbr """)
AyDmp A
End Sub

Sub Lis_Md_Mth(Optional MthNmPatn$ = ".", Optional Mdy0$)
AyDmp AyAddPfx(MdMthNy(CurMd, MthNmPatn, Mdy0:=Mdy0), CurPjNm & ".")
End Sub

Sub Lis_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$)
Lis_Vbe_Mth MthNmPatn, MdNmPatn, Mdy
End Sub

Sub Lis_Pj()
Dim A$()
    A = CurVbe_PjNy
    A = AyAddPfx(A, "Shw_Pj """)
AyDmp A
End Sub

Sub Lis_Pj_Mth(Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".", Optional Mdy0$)
Dim A$()
    A = PjMthNy(CurPj, MthNmPatn:=MthNmPatn, MbrNmPatn:=MbrNmPatn, Mdy0:=Mdy0)
    A = AySrt(A)
    A = AyAddPfx(A, "Shw_Mth """)
AyDmp A
End Sub

Sub Lis_Vbe_DupMthFNy(Optional IsSamMthBdyOnly As Boolean)
Dim A$(): A = VbeDupMthFNy(CurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
Dim A1$(): A1 = AyDblQuote(A)
AyDmp AyAddPfxSfx(A1, "Shw_Mth ", ",IsMthPjMdNm:=True")
End Sub

Sub Lis_Vbe_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$)
Dim A$()
    A = CurVbe_MthNy(MthNmPatn, MdNmPatn, Mdy)
    A = AySrt(A)
AyDmp A
End Sub

Sub Mov_Fun(Optional MthDNm0$)
'Mov Fun to its proper-module
'Fun here means Public-Prp/Sub/Fun, in a Md, not class
'         or    Private-Sub ZZ_xxx, in a Md, not class
'proper-module means, M_Xxx where Xxx is function-MdPfx
'MdPfx-of-a-fun is a Pfx of a funNm which is used to give a proper-module-nm of M_Xxx
Dim M As Mth: Set M = DftMth(MthDNm0)
If MdCmpTy(M.Md) <> vbext_ct_StdModule Then
    Debug.Print FmtQQ("Mov_Fun: CurMth(?) in not in StdMd", MthDNm(M))
    Exit Sub
End If
If Not IsPfx(M.Nm, "ZZ_") Then
    If Not MthIsPub(M) Then
        Debug.Print FmtQQ("Mov_Fun: CurMth(?) is not public", MthDNm(M))
        Exit Sub
    End If
End If
MthMovToProperMd M
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
Ny = MdMthNy_OfInproper(CurMd)
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
    MdMov_ToPj CvMd(I), P
Next
Cls_Win
End Sub

Sub Mov_Md_ToPj(ToPjNm$)
If CurPjNm = ToPjNm Then
    Debug.Print FmtQQ("Mov_Md: ToPjNm(?) cannot be CurPjNm", ToPjNm)
    Exit Sub
End If
MdMov_ToPj CurMd, Pj(ToPjNm)
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
MthRmk CurMth
End Sub

Sub Sav_Pj()
PjSav CurPj
End Sub

Sub Sav_Vbe()
Dim I
For Each I In CurVbe_PjAy
    PjSav CvPj(I)
Next
End Sub

Function Shw_Pj_SrtRptWb(Optional PjNm$) As Workbook
PjSrtRptWb DftPj(PjNm), Vis:=True
End Function

Sub Srt_F_Tool()
Dim M As CodeModule: Set M = Md("QTool.F_Tool")
Dim Src$(): Src = MdSrc(M)
Dim Cxt$: Cxt = SrcSrtedLines(Src)
If Cxt = Join(Src, vbCrLf) Then
    Debug.Print "Md(F_Tool) is already sorted"
Else
    MdRpl_Cxt M, Cxt
End If
End Sub

Sub Srt_G_Tool()
Dim M As CodeModule: Set M = Md("QTool.G_Tool")
Dim Src$(): Src = MdSrc(M)
Dim Cxt$: Cxt = SrcSrtedLines(Src)
If Cxt = Join(Src, vbCrLf) Then
    Debug.Print "Md(F_Tool) is alread sorted"
Else
    MdRpl_Cxt M, Cxt
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
End Sub
