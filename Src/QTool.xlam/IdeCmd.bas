Attribute VB_Name = "IdeCmd"
Option Explicit
Function MovPjMthLisAy(Pj$, MthPatn$, ToMd$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$, Optional MdExl$) As String()
Dim FNm$(), Fn, Mth$, Md$, MdExl1$
MdExl1 = ToMd & IIf(MdExl <> "", " ", "") & MdExl
FNm = PjMthFNyWh(CurVbePj(Pj), WhMdMth(MthPatn, MthExl, WhMdy, WhKd, MdPatn, MdExl:=MdExl1, WhCmpTy:="Std"))
For Each Fn In AyNz(FNm)
    Mth = Brk(Fn, ":").S1
    Md = Pj & "." & ToMd
    Push MovPjMthLisAy, FmtQQ("MthMov DMth(""?""), Md(""?"") '?", Mth, Md, Fn)
Next
End Function

Sub MovPjMthGen(Pj$, Patn$, ToMd$, Pfx$)
CdAyGen MovPjMthLisAy(Pj$, Patn, ToMd$), FmtQQ("Mov_Pj_?_Pfx_?_To_?", Pj, Pfx, ToMd)
End Sub

Sub LisMovPjMth(Pj$, MthPatn$, ToMd$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$, Optional MdExl$)
Dim Ay$()
Ay = MovPjMthLisAy(Pj, MthPatn$, ToMd$, MthExl$, WhMdy$, WhKd$, MdPatn$, MdExl$)
D AyFmt(Ay, ") ' :")
End Sub

Sub Exe(MthPfx$, ToMd$)
Const P$ = "QTool"
Run FmtQQ("Mov_Pj_?_Pfx_?_To_?", P, MthPfx, ToMd)
End Sub

Sub Gen(Pfx$, ToMd$)
Dim Patn$
Patn = Replace("^(ZZ_?|Z_?|?)", "?", Pfx)
MovPjMthGen "QTool", Patn, ToMd, Pfx
End Sub
Sub ActPj(PjNm$)
Set CurVbe.ActiveVBProject = Pj(PjNm)
End Sub
Sub Lis(Pfx$, ToMd$)
Const P$ = "QTool"
PjEnsMd Pj(P), ToMd
ActPj "QToolTmp"
Dim Patn$
Patn = Replace("^(ZZ_?|Z_?|?)", "?", Pfx)
LisMovPjMth P, Patn, ToMd
End Sub
Sub LisA()
LisMth "^AA"
End Sub
Sub LisCurMth()
Debug.Print MthLines(CurMth)
End Sub
Sub LisMd(Optional Patn$, Optional Exl$)
Dim A$()
    A = PjCmpNyWh(CurPj, "Md", Patn, Exl)
    A = AySrt(A)
    A = AyAddPfx(A, "ShwMbr """)
D A
End Sub
Sub LisMdMthPfx()
D AySrt(MdMthPfx(CurMd))
End Sub
Sub LisMdMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$)
Dim Ny$(), M As WhMth
M = WhMth(MthPatn, MthExl, WhMdy, WhKd)
Ny = MdMthNyWh(CurMd, M)
D AyAddPfx(Ny, CurPjNm & ".")
End Sub
Function WhPjMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$, Optional MdExl$, Optional WhCmpTy$, Optional PjPatn$, Optional PjExl$) As WhPjMth
With WhPjMth
    .Pj = WhNm(PjPatn, PjExl)
    .MdMth = WhMdMth(MthPatn, MthExl, WhMdy, WhKd, MdPatn, MdExl, WhCmpTy)
End With
End Function
Sub LisMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$, Optional MdExl$, Optional WhCmpTy$, Optional PjPatn$, Optional PjExl$)
D VbeMthNyWh(CurVbe, WhPjMth(MthPatn, MthExl, WhMdy, WhKd, MdPatn, MdExl, WhCmpTy, PjPatn, PjExl))
End Sub
Sub MovPjMth(MthPatn$, ToMd$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$, Optional MdExl$)
CdAyRun MovPjMthLisAy(MthPatn$, ToMd$, MthExl$, WhMdy$, WhKd$, MdPatn$, MdExl$)
End Sub

Function LinBrkssDr(Lin, BrkssAy$()) As String()
Dim Brk, P%, L$
L = Lin
For Each Brk In BrkssAy
    P = InStr(L, Brk)
    If P = 0 Then Exit For
    Push LinBrkssDr, Left(L, P - 1)
    L = Mid(L, P)
Next
Push LinBrkssDr, L
End Function

Sub LisPj()
Dim A$()
    A = VbePjNy(CurVbe)
    D AyAddPfx(A, "ShwPj """)
D A
End Sub
Sub LisPjDupMth(Optional IsSamMthBdyOnly As Boolean)
D PjDupMth(CurPj, IsSamMthBdyOnly:=IsSamMthBdyOnly)
End Sub
Sub LisPjFunPfx()
D AySrt(PjFunPfxAy(CurPj))
End Sub
Sub LisVbeMthPfx()
D AySrt(VbeMthPfx(CurVbe))
End Sub
Sub LisVbeDupMth()
Stop '
'DrsBrw VbeDupMthDryWh(CurVbe)
End Sub
Sub LisVbeMth(Optional MthPatn$, Optional MdPatn$, Optional Mdy$)
Dim A$()
Stop '
'    A = VbeMthNyWh(CurVbe, CvPatn(MthPatn), MdPatn, Mdy)
    A = AySrt(A)
D AyAddPfx(A, "Shw """)
End Sub
Sub AddMd(Nm$)
PjAddCmp CurPj, Nm, vbext_ComponentType.vbext_ct_StdModule
End Sub
Sub AddCls(Nm$)
PjAddCmp CurPj, Nm, vbext_ComponentType.vbext_ct_ClassModule
End Sub
