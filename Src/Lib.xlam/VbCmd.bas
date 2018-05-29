Attribute VB_Name = "VbCmd"
Option Explicit
Enum eCmpTySelection
    eMdOnly = 1
    eClsOnly = 2
    eBothMdAndCls = 3
End Enum

Sub AddCls(Nm$)
CurPjx.CrtMd Nm, vbext_ct_ClassModule
'MdGo Nm
End Sub

Sub AddFun(FunNm$, Optional MdNm$)
'Des: Add Empty-Fun-Mth to CurMd
MdAppLines DftMdByMdNm(MdNm), FmtQQVBar("Function ?()|End Function", FunNm)
MdMth_Go DftMdByMdNm(MdNm), FunNm
End Sub

Sub AddMd(Nm$)
'PjCrtMd CurPj, Nm, vbext_ct_StdModule
'MdGo Nm
End Sub

Sub AddSub(SubNm$, Optional MdNm$)
'Des: Add Sub-Mth to CurMd
'MdAppLines FmtQQVBar("Sub ?()|End Sub", SubNm)
'MthGo SubNm, DftMdByMdNm(MdNm)
End Sub

Sub DltMd()
If MsgBox(FmtQQ("Delete this Md[?]", CurMdNm), vbYesNo + vbDefaultButton2) = vbYes Then

End If
End Sub

Sub GG()
GoMd "XlsLoFmt"
End Sub

Sub GoCls(ClsNm$)
MdGo IdeMd.Md(ClsNm)
End Sub

Sub GoMd(MdNm$)
'MdGo Md(MdNm)
End Sub

Sub LisMdMth(Optional MthNmPatn$ = ".")
AyDmp AyAddPfx(MdMthNy(CurMd, MthNmPatn), CurMdNm & ".")
End Sub

Sub LisMth(Optional MthNmPatn$ = ".")
LisMdMth MthNmPatn
End Sub

Sub LisPjMth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
AyDmp CurPjx.MthNy(, MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
End Sub

Sub RenMd(NewNm$)
CurMd.Name = NewNm
End Sub

Sub SrtMd(Optional MdNm0$)
MdSrt DftMdByMdNm(MdNm0)
End Sub

Sub SrtPj()
PjSrt CurPj
End Sub
