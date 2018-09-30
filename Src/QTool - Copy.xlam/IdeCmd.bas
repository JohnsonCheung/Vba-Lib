Attribute VB_Name = "IdeCmd"
Option Explicit
Sub LisA()
LisMth WhPjMth(WhNm("^AA"))
End Sub
Sub LisCurMth()
Debug.Print MthLines(CurMth)
End Sub
Sub LisMd(Optional Patn$, Optional Exl$)
Dim A$()
    A = PjCmpNy(CurPj, WhMd("Std", WhNm(Patn, Exl)))
    A = AySrt(A)
    A = AyAddPfx(A, "ShwMbr """)
D A
End Sub
Function MdMthPfx(A As CodeModule) As String()
End Function
Sub LisMdMthPfx()
D AySrt(MdMthPfx(CurMd))
End Sub
Sub LisMdMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$)
Dim Ny$(), M As WhMth
M = WhMth(WhMdy, WhKd, WhNm(MthPatn, MthExl))
Ny = MdMthNy(CurMd, M)
D AyAddPfx(Ny, CurPjNm & ".")
End Sub

Function WhPjMth(Optional Pj As WhNm, Optional MdMth As WhMdMth) As WhPjMth
Set WhPjMth = New WhPjMth
With WhPjMth
    Set .Pj = Pj
    Set .MdMth = MdMth
End With
End Function

Sub LisMth(Optional A As WhPjMth)
D VbeMthNy(CurVbe, A)
End Sub

Sub SetActPj(PjNm$)
Set CurVbe.ActiveVBProject = Pj(PjNm)
End Sub

Function CurVbePj(A$) As VBProject
Set CurVbePj = CurVbe.VBProjects(A)
End Function

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

Sub LisPjFunPfx()
D AySrt(PjFunPfxAy(CurPj))
End Sub
Sub LisVbeMthPfx()
D AySrt(VbeMthPfx(CurVbe))
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
