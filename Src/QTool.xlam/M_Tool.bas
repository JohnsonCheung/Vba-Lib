Attribute VB_Name = "M_Tool"
Option Explicit

Sub AAA()
Rmk_Mth
End Sub

Sub Brw_DupMdNm()
ZAyBrw ZCurVbe_DupMdNy
End Sub

Sub Add_Cls(Nm$)
ZPj_Add_Mbr ZCurPj, Nm, vbext_ct_ClassModule
End Sub

Sub Add_CurVbe_QToolRf()
Dim I, P As VBProject
For Each I In ZCurVbe_PjAy
    Set P = I
    ZPj_AddRf P, "QTool"
Next
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

Sub Brw_Md_MthKy()
ZAyBrw ZMd_MthKy(ZCurMd, IsSngLinFmt:=True)
End Sub

Sub Brw_Pj_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
ZAyBrw ZCurPj_MthNy(MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
End Sub

Sub Brw_Pj_MthKy()
ZAyBrw ZPj_MthKy(ZCurPj, IsSngLinFmt:=True)
End Sub

Sub Brw_Pj_SrcPth()
ZPj_SrcPthBrw ZCurPj
End Sub

Sub Brw_Pj_SrtRpt()
ZAyBrw ZPj_SrtRptLy(ZCurPj)
End Sub

Sub Brw_Vbe_MthKy()
ZAyBrw ZVbe_MthKy(ZCurVbe, IsSngLinFmt:=True)
End Sub

Sub Brw_Vbe_SrtRpt()
ZAyBrw ZVbe_SrtRptLy(ZCurVbe)
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
ZAyDo ZCurVbe_PjAy, "ZPj_Compile"
End Sub

Sub Cpy_Mbr(FmPjMbrDotNm$)
ZMd_Cpy_ToPj ZMd(FmPjMbrDotNm), ZCurPj
End Sub

Sub Dlt_Md()
If MsgBox(ZFmtQQ("Delete this Md[?]", ZCurMdNm), vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
ZCurPj.VBComponents.Remove ZCurCmp
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

Sub Go_Mth(PjMdMthDotNm$)
Dim Md As CodeModule
Dim MthNm$
ZPjMdMthDotNm_BrkAsg PjMdMthDotNm, _
    Md, MthNm
Dim L As LCCOpt
    L = ZMdMth_LCCOpt(Md, MthNm)
ZMd_GoLCCOpt Md, L
End Sub

Sub Go_Pj(PjNm$)
ZPj_Go ZPj(PjNm)
End Sub

Sub Lis_Md()
Dim A$()
    A = ZCurPj_MbrNy
    A = ZAySrt(A)
    A = ZAyAddPfx(A, "Go_Mbr """)
ZAyDmp A
End Sub

Sub Lis_Md_Mth(Optional MthNmPatn$ = ".")
ZAyDmp ZAyAddPfx(ZCurMd_MthNy(MthNmPatn), ZCurMdNm & ".")
End Sub

Sub Lis_Mth(Optional MthNmPatn$ = ".")
Lis_Vbe_Mth MthNmPatn
End Sub

Sub Lis_Pj()
Dim A$()
    A = ZCurVbe_PjNy
    A = ZAyAddPfx(A, "Go_Pj """)
ZAyDmp A
End Sub

Sub Lis_Pj_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
Dim A$()
    A = ZCurPj_MthNy(MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
    A = ZAySrt(A)
    A = ZAyAddPfx(A, "Go_Mth """)
ZAyDmp A
End Sub

Sub Lis_Vbe_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
Dim A$()
    A = ZCurVbe_MthNy(MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
    A = ZAySrt(A)
ZAyDmp A
End Sub

Sub Mov_MdLik_ToPj(MdLikNm$, ToPjNm$)
Dim Ay() As CodeModule: Ay = ZCurPj_MbrAyLik(MdLikNm)
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
If ZCurPj_HasCmp(NewNm) Then
    MsgBox ZFmtQQ("Md(?) exists in CurPj(?).  Cannot rename.", NewNm, ZCurPjNm), , "M_A:RenMd"
    Exit Sub
End If
ZCurMd.Name = NewNm
End Sub

Sub Rmk_All()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In ZMbrAy
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
ZMdMth_Rmk_Bdy ZCurMd, ZCurMthNm
End Sub

Function Shw_CurPj_SrtRptWb(Optional Vis As Boolean) As Workbook
ZPj_SrtRptWb ZCurPj, Vis
End Function

Sub Srt_M_Tool()
Dim P As VBProject
Dim Md As CodeModule
Dim Src$()
Dim Cxt$
Set P = ZPj("QTool")
Set Md = ZPj_Md(P, "M_Tool")
Src = ZMd_Src(Md)
Cxt = ZSrc_SrtedLines(Src)
ZPj_Ens_Md P, "M_Tool1", Cxt
End Sub

Sub Srt_Md()
ZMd_Srt ZCurMd
End Sub

Sub Srt_Pj()
ZPj_Srt ZCurPj
End Sub

Sub UnRmk_All()
Dim I, Md As CodeModule
Dim NUnRmk%, Skip%
For Each I In ZMbrAy
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
