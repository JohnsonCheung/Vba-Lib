Attribute VB_Name = "F_Tool"
Option Explicit

Sub Add_Cls(Nm$)
ZPj_Add_Mbr ZCurPj, Nm, vbext_ct_ClassModule
End Sub

Sub Add_Fun(FunNm$)
ZAdd_Fun_or_Sub FunNm, IsFun:=True
End Sub

Sub Add_Md(Nm$)
ZPj_Add_Mbr ZCurPj, Nm, vbext_ct_StdModule
End Sub

Sub Add_Sub(SubNm$)
ZAdd_Fun_or_Sub SubNm, IsFun:=False
End Sub

Sub Add_VbeRf_QTool()
Dim I, P As VBProject
For Each I In ZCurVbe_PjAy
    Set P = I
    ZPj_AddRf P, "QTool"
Next
End Sub

Sub Brw_DupMdNm()
AyBrw ZCurVbe_DupMdNy
End Sub

Sub Brw_InproperMth()
Brw_Pj_InproperMth
End Sub

Sub Brw_Md_InproperMth()
AyBrw ZMd_MthNy_OfInproper(ZCurMd)
End Sub

Sub Brw_Md_Mth()
ZS1S2Ay_Brw ZMd_MthS1S2Ay(ZCurMd)
End Sub

Sub Brw_Md_MthKy()
AyBrw ZMd_MthKy(ZCurMd, IsSngLinFmt:=True)
End Sub

Sub Brw_Md_MthNm(Optional MthNmPatn$, Optional Mdy$)
AyBrw ZMd_MthNy(ZCurMd, MthNmPatn, IsNoMdNmPfx:=True, Mdy:=Mdy)
End Sub

Sub Brw_Pj_FFunNy()
AyBrw ZPj_FFunNy(ZCurPj)
End Sub

Sub Brw_Pj_InproperMth()
AyBrw ZPj_MthNy_OfInproper(ZCurPj)
End Sub

Sub Brw_Pj_Mth(Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".")
AyBrw ZPj_MthNy(ZCurPj, MthNmPatn:=MthNmPatn, MbrNmPatn:=MbrNmPatn)
End Sub

Sub Brw_Pj_MthKy()
AyBrw ZPj_MthKy(ZCurPj, IsSngLinFmt:=True)
End Sub

Sub Brw_Pj_SrtRpt()
AyBrw ZPj_SrtRptLy(ZCurPj)
End Sub

Sub Brw_Vbe_DupFFunDrs(Optional IsSamMthBdyOnly As Boolean)
ZWsVis ZDrsWs(ZVbe_DupFFunDrs(ZCurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly))
End Sub

Sub Brw_Vbe_DupFFunNy(Optional IsSamMthBdyOnly As Boolean)
AyBrw ZVbe_DupFFunNy(ZCurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Brw_Vbe_DupFun()
AyBrw ZVbe_DupFunLy(ZCurVbe)
End Sub

Sub Brw_Vbe_FFunNy()
AyBrw ZVbe_FFunNy(ZCurVbe)
End Sub

Sub Brw_Vbe_InproperMth()
AyBrw ZVbe_MthNy_OfInproper(ZCurVbe)
End Sub

Sub Brw_Vbe_MthKy()
AyBrw ZVbe_MthKy(ZCurVbe, IsSngLinFmt:=True)
End Sub

Sub Brw_Vbe_SrcPth()
ZVbe_SrcPthBrw ZCurVbe
End Sub

Sub Brw_Vbe_SrtRpt()
AyBrw ZVbe_SrtRptLy(ZCurVbe)
End Sub

Sub Cls_Win()
Dim W As VBIDE.Window
For Each W In ZCurVbe.Windows
    W.Close
Next
End Sub

Sub Compile_Pj()
ZPj_Compile ZCurPj
End Sub

Sub Compile_Vbe()
AyDo ZCurVbe_PjAy, "ZPj_Compile"
End Sub

Sub Cpy_Mbr(FmPjMbrDotNm$)
ZMd_Cpy_ToPj ZMd(FmPjMbrDotNm), ZCurPj
End Sub

Sub Cpy_Md_ToPj(ToPjNm$)
ZMd_Cpy_ToPj ZCurMd, ZPj(ToPjNm)
End Sub

Sub Dlt_Md()
If MsgBox(ZFmtQQ("Delete this Md[?]", ZCurMdNm), vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
ZCurPj.VBComponents.Remove ZCurCmp
End Sub

Sub Dmp_CurMth()
Debug.Print ZMth_Lines(ZCurMth)
End Sub

Sub Dmp_Md_InproperMth()
AyDmp ZMd_MthNy_OfInproper(ZCurMd)
End Sub

Sub Dmp_Pj_DupFFunNy(Optional IsSamMthBdyOnly As Boolean)
AyDmp ZPj_DupFFunNy(ZCurPj, IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Dmp_Pj_InproperMth(Optional PjNm$)
AyDmp ZPj_MthNy_OfInproper(ZDft_Pj(PjNm))
End Sub

Sub Dmp_Vbe_DupFFunNy(Optional IsSamMthBdyOnly As Boolean)
AyDmp ZVbe_DupFFunNy(ZCurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub

Sub Dmp_Vbe_InproperMth()
AyDmp ZVbe_MthNy_OfInproper(ZCurVbe)
End Sub

Sub Export_Pj()
ZPj_Export ZCurPj
End Sub

Sub Export_Vbe()
ZVbe_Export ZCurVbe
End Sub

Sub Gen_Md_TstSub()
ZMd_Gen_TstSub ZCurMd
End Sub

Sub Gen_Pj_ConstructorMd()
Stop '
End Sub

Sub Gen_Pj_TstClass()
ZPj_Gen_TstClass ZCurPj
End Sub

Sub Gen_Pj_TstSub()
ZPj_Gen_TstSub ZCurPj
End Sub

Sub Gen_Vbe_TstClass()
End Sub

Sub Go_Mbr(PjMbrDotNm$)
Dim E As Either
E = ZPjMbrDotNm_Either(PjMbrDotNm)
If E.IsLeft Then
    ZMd_Go ZMd(E.Left)
    Exit Sub
End If
Dim Ny$()
    Ny = E.Right
If ZSz(Ny) = 0 Then
    Debug.Print PjMbrDotNm; "<-- No such module"
    Stop '
    Exit Sub
End If
Dim I
For Each I In Ny
    Debug.Print "Go_Mbr """; I; "."; PjMbrDotNm
Next
End Sub

Sub Go_Mth(MthDNm$)
Dim M As Mth
Set M = ZMthDNm_Mth(MthDNm)
ZMd_GoLCCOpt M.Md, ZMth_LCCOpt(M)
End Sub

Sub Go_Pj(PjNm$)
ZPj_Go ZPj(PjNm)
End Sub

Sub Lis_Md()
Dim A$()
    A = ZPj_MbrNy(ZCurPj)
    A = AySrt(A)
    A = AyAddPfx(A, "Go_Mbr """)
AyDmp A
End Sub

Sub Lis_Md_Mth(Optional MthNmPatn$ = ".", Optional Mdy$)
AyDmp AyAddPfx(ZMd_MthNy(ZCurMd, MthNmPatn, Mdy:=Mdy), ZCurPjNm & ".")
End Sub

Sub Lis_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$)
Lis_Vbe_Mth MthNmPatn, MdNmPatn, Mdy
End Sub

Sub Lis_Pj()
Dim A$()
    A = ZCurVbe_PjNy
    A = AyAddPfx(A, "Go_Pj """)
AyDmp A
End Sub

Sub Lis_Pj_Mth(Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".", Optional Mdy$)
Dim A$()
    A = ZPj_MthNy(ZCurPj, MthNmPatn:=MthNmPatn, MbrNmPatn:=MbrNmPatn, Mdy:=Mdy)
    A = AySrt(A)
    A = AyAddPfx(A, "Go_Mth """)
AyDmp A
End Sub

Sub Lis_Vbe_DupFFunNy(Optional IsSamMthBdyOnly As Boolean)
Dim A$(): A = ZVbe_DupFFunNy(ZCurVbe, ExclPjNy0:="QLib", IsSamMthBdyOnly:=IsSamMthBdyOnly)
Dim A1$(): A1 = AyDblQuote(A)
AyDmp AyAddPfxSfx(A1, "Go_Mth ", ",IsMthPjMdNm:=True")
End Sub

Sub Lis_Vbe_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$)
Dim A$()
    A = ZCurVbe_MthNy(MthNmPatn, MdNmPatn, Mdy)
    A = AySrt(A)
AyDmp A
End Sub

Sub Mov_Fun()
'Mov Fun to its proper-module
'Fun here means Public-Prp/Sub/Fun, in a Md, not class
'proper-module means, M_Xxx where Xxx is function-MdPfx
'MdPfx-of-a-fun is a Pfx of a funNm which is used to give a proper-module-nm of M_Xxx
With ZCurPubPSFunOpt
    Stop
    If Not .Som Then
        Debug.Print "Mov_Fun: No Cur Public Prp-Sub-Fun.  Cannot Mov"
        Exit Sub
    End If
    ZMth_Mov .Mth, ZMth_ProperMd(.Mth)
End With
End Sub

Sub Mov_Fun_ToProperMd()
'Move all Inproper-Fun in CurMd to its proper module in same Pj
'If non-exist-inproper-module will be created
'If a Fun in a module of name of format M_XXX,
'   if the Fun-name-pfx is not XXX, => it is inproper-fun
'else
'   => it is proper-fun
Dim I, M As CodeModule, Ny$()
Set M = ZCurMd
Ny = ZMd_MthNy_OfInproper(ZCurMd)
If ZSz(Ny) = 0 Then Exit Sub
Dim N
Dim Mth As Mth
Set Mth.Md = M
For Each N In Ny
    Mth.Nm = N
    ZMth_Mov_ToProperMd Mth
Next
End Sub

Sub Mov_MbrPatn_ToPj(MbrNmPatn$, ToPjNm$)
Dim Ay() As CodeModule: Ay = ZPj_MbrAy(ZCurPj, MbrNmPatn)
If ZSz(Ay) = 0 Then Exit Sub
Dim I, P As VBProject
Set P = ZPj(ToPjNm)
For Each I In Ay
    ZMd_Mov_ToPj ZCvMd(I), P
Next
ZClsWinExcept_Module_A_1
End Sub


Sub Mov_Md_ToPj(ToPjNm$)
If ZCurPjNm = ToPjNm Then
    Debug.Print ZFmtQQ("Mov_Md: ToPjNm(?) cannot be CurPjNm", ToPjNm)
    Exit Sub
End If
ZMd_Mov_ToPj ZCurMd, ZPj(ToPjNm)
End Sub

Sub Ren_Md(NewNm$)
If ZPj_HasCmp(ZCurPj, NewNm) Then
    MsgBox ZFmtQQ("Md(?) exists in CurPj(?).  Cannot rename.", NewNm, ZCurPjNm), , "M_A:RenMd"
    Exit Sub
End If
ZCurMd.Name = NewNm
End Sub

Sub Rmk_All()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In ZPj_MbrAy(ZCurPj)
    Set Md = I
    If ZMd_Rmk(Md) Then
        NRmk = NRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub

Sub Rmk_Mth()
ZMth_Rmk_Bdy ZCurMth
End Sub

Sub Sav_Pj()
ZPj_Sav ZCurPj
End Sub

Sub Sav_Vbe()
Dim I
For Each I In ZCurVbe_PjAy
    ZPj_Sav ZCvPj(I)
Next
End Sub

Function Shw_Pj_SrtRptWb(Optional PjNm$) As Workbook
ZPj_SrtRptWb ZDft_Pj(PjNm), Vis:=True
End Function

Sub Srt_F_Tool()
Dim P As VBProject
Dim Md As CodeModule
Dim Src$()
Dim Cxt$
Set P = ZPj("QTool")
Set Md = ZPj_Md(P, "F_Tool")
Src = ZMd_Src(Md)
Cxt = ZSrc_SrtedLines(Src)
ZMd_Ens_Cxt Md, Cxt
End Sub

Sub Srt_Md()
ZMd_Srt ZCurMd
End Sub

Sub Srt_Pj()
ZPj_Srt ZCurPj
End Sub

Sub Srt_Vbe()
ZVbe_Srt ZCurVbe
End Sub

Sub UnRmk_All()
Dim I, Md As CodeModule
Dim NUnRmk%, Skip%
For Each I In ZPj_MbrAy(ZCurPj)
    Set Md = I
    If ZMd_UnRmk(Md) Then
        NUnRmk = NUnRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NUnRmk"; NUnRmk
Debug.Print "SKip"; Skip
End Sub

Sub UnRmk_Mth()
ZMth_UnRmk_Bdy ZCurMth
End Sub
