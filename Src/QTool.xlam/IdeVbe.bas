Attribute VB_Name = "IdeVbe"
Option Explicit
Sub VbeEnsZZDashPubMthAsPrivate(A As Vbe)
AyDo VbePjAy(A), "PjEnsZZDashPubMthAsPrivate"
End Sub
Function VbeMthLinDry(A As Vbe) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushAy VbeMthLinDry, PjMthLinDry(CvPj(P))
Next
End Function
Function VbeMthLinDryWP(A As Vbe) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushIAy VbeMthLinDryWP, PjMthLinDryWP(CvPj(P))
Next
End Function
Function VbeAyMthWs(A() As Vbe) As Worksheet
Set VbeAyMthWs = DrsWs(VbeAyMthDrs(A))
End Function
Function VbeAyMthDrs(A() As Vbe) As Drs
Dim I, R%, M As Drs
For Each I In AyNz(A)
    Set M = DrsInsCol(VbeMthDrs(CvVbe(I)), "Vbe", R)
    If R = 0 Then
        Set VbeAyMthDrs = M
    Else
        Stop
        PushDrs VbeAyMthDrs, M
        Stop
    End If
    R = R + 1
    Debug.Print R; "<=== VbeAyMthDrs"
Next
End Function
Function VbeMthDrsWh(A As Vbe, B As WhMth, Optional C As MthBrkOpt) As Drs
Set VbeMthDrsWh = Drs(MthBrkOptFny(C), VbeMthDryWh(A, B, C))
End Function
Function VbeMthDrs(A As Vbe) As Drs
Dim O As Drs, O1 As Drs, O2 As Drs
Set O = Drs("Pj Md Mdy Ty Nm Lines", VbeMthDry(A))
Set O1 = DrsAddValIdCol(O, "Nm")
Set O2 = DrsAddValIdCol(O1, "Lines")
Set VbeMthDrs = O2
End Function
Function VbeMthDryWh(A As Vbe, B As WhMth, Optional C As MthBrkOpt) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushAy VbeMthDryWh, PjMthDryWh(CvPj(P), B, C)
Next
End Function
Function VbeMthFny() As String()
VbeMthFny = ApSy("Pj", "Md", "Mdy", "Ty", "Nm", "Lines")
End Function
Function VbeMthDry(A As Vbe) As Variant()
Dim P
For Each P In AyNz(VbePjAy(A))
    PushAy VbeMthDry, PjMthDry(CvPj(P))
Next
End Function
Function VbeMthWb(A As Vbe) As Workbook
Set VbeMthWb = WbVis(WbSavAs(WsWb(VbeMthWs(A)), VbeMthFx))
End Function
Function VbeMthWs(A As Vbe) As Worksheet
Set VbeMthWs = DrsWs(VbeMthDrs(A))
End Function
Function VbeMthFx$()
VbeMthFx = FfnNxt(CurPjPth & "VbeMth.xlsx")
End Function
Function VbeHasPj(A As Vbe, PjNm) As Boolean
VbeHasPj = ItrHasNm(A.VBProjects, PjNm)
End Function
Function VbeMthMdDNm$(A As Vbe, MthNm)
Dim Pj As VBProject, P, M, Md As CodeModule
For Each P In AyNz(VbePjAy(A))
    Set Pj = P
    For Each M In PjMdAy(Pj)
        Set Md = M
        If MdHasMth(Md, MthNm) Then VbeMthMdDNm = MdDNm(Md) & "." & MthNm: Exit Function
    Next
Next
End Function
Function VbeMthMdDNy(A As Vbe, MthNm) As String()
Dim Pj As VBProject, P, M, Md As CodeModule
For Each P In AyNz(VbePjAy(A))
    Set Pj = P
    For Each M In PjMdAy(Pj)
        Set Md = M
        If MdHasMth(Md, MthNm) Then Push VbeMthMdDNy, MdDNm(Md) & "." & MthNm
    Next
Next
End Function
Function VbePjNy(A As Vbe, Optional Patn$, Optional Exl$) As String()
VbePjNy = AyWhPatnExl(ItrNy(A.VBProjects), Patn, Exl)
End Function
Function VbeDupMthCmpLy(A As Vbe, B As WhPjMth, Optional InclSam As Boolean) As String()
Stop
Dim N$(): 'N = VbeDupMthFNm(A, B)
Dim Ay(): Ay = DupMthFNy_GpAy(N)
Dim O$(), J%
Push O, FmtQQ("Total ? dup function.  ? of them has mth-lines are same", Sz(Ay), DupMthFNyGpAy_AllSameCnt(Ay))
Dim Cnt%, Sam%
For J = 0 To UB(Ay)
    PushAy O, DupMthFNyGp_CmpLy(Ay(J), Cnt, Sam, InclSam:=InclSam)
Next
VbeDupMthCmpLy = O
End Function
Function VbeDupMthDrs(A As Vbe, B As WhPjMth, Optional IsSamMthBdyOnly As Boolean, Optional IsNoSrt As Boolean) As Drs
Dim Fny$(), Dry()
Fny = SplitSsl("Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src")
Dry = VbeDupMthDryWh(A, B, IsSamMthBdyOnly:=IsSamMthBdyOnly)
Set VbeDupMthDrs = Drs(Fny, Dry)
End Function
Function VbeDupMthDry(A As Vbe) As Variant()
Dim B(): B = VbeMthDry(A)
Dim Ny$(): Ny = DryStrCol(B, 2)
Dim N1$(): N1 = AyWhDup(Ny)
    N1 = DupMthFNy_SamMthBdyFunFNy(N1, A)
Dim GpAy()
    GpAy = DupMthFNy_GpAy(N1)
    If Sz(GpAy) = 0 Then Exit Function
Dim O()
    Dim Gp
    For Each Gp In GpAy
        PushAy O, DupMthFNyGp_Dry(CvSy(Gp))
    Next
VbeDupMthDry = O
End Function
Function VbeDupMthDryWh(A As Vbe, B As WhPjMth, Optional IsSamMthBdyOnly As Boolean) As Variant()
Dim N$(): 'N = VbeFunFNm(A)
Dim N1$(): ' N1 = MthNyWhDup(N)
    If IsSamMthBdyOnly Then
        N1 = DupMthFNy_SamMthBdyFunFNy(N1, A)
    End If
Dim GpAy()
    GpAy = DupMthFNy_GpAy(N1)
    If Sz(GpAy) = 0 Then Exit Function
Dim O()
    Dim Gp
    For Each Gp In GpAy
        PushAy O, DupMthFNyGp_Dry(CvSy(Gp))
    Next
VbeDupMthDryWh = O
End Function
Function VbePj(A As Vbe, Pj$) As VBProject
Set VbePj = A.VBProjects(Pj)
End Function
Sub VbePjMdFmtBrw(A As Vbe)
Brw VbePjMdFmt(A)
End Sub
Function VbePjMdFmt(A As Vbe) As String()
VbePjMdFmt = DryFmtss(VbePjMdDry(A))
End Function
Function VbePjMdDry(A As Vbe) As Variant()
Dim O(), P, C, PNm$, Pj As VBProject
For Each P In VbePjAy(A)
    Set Pj = P
    PNm = PjNm(Pj)
    For Each C In PjCmpAy(Pj)
        Push O, Array(PNm, CvCmp(C).Name)
    Next
Next
VbePjMdDry = O
End Function
Function VbeDupMdNy(A As Vbe) As String()
VbeDupMdNy = DryFmtss(DryWhDup(VbePjMdDry(A)))
End Function
Function VbeFstQPj(A As Vbe) As VBProject
Dim I
For Each I In A.VBProjects
    If FstChr(CvPj(I).Name) = "Q" Then
        Set VbeFstQPj = I
        Exit Function
    End If
Next
End Function
Function VbePjNyWhMd(A As Vbe, MdPatn$) As String()
Dim I, Re As New RegExp
Re.Pattern = MdPatn
For Each I In VbePjAy(A)
    If PjHasCmpWhRe(CvPj(I), Re) Then
        Push VbePjNyWhMd, CvPj(I).Name
    End If
Next
End Function
Function VbeMthKy(A As Vbe, Optional IsWrap As Boolean) As String()
Dim O$(), I
For Each I In VbePjAy(A)
    PushAy O, PjMthKy(CvPj(I), IsWrap)
Next
VbeMthKy = O
End Function
Function VbeMthNy(A As Vbe) As String()
Dim I
For Each I In AyNz(VbePjAy(A))
    PushAy VbeMthNy, PjMthNy(CvPj(I))
Next
End Function
Function VbeMthNyWh(A As Vbe, B As WhPjMth) As String()
Dim I
For Each I In AyNz(VbePjAyWh(A, B.Pj))
    PushAy VbeMthNyWh, PjMthNyWh(CvPj(I), B.MdMth)
Next
End Function
Function VbePjAy(A As Vbe) As VBProject()
VbePjAy = ItrAyInto(A.VBProjects, VbePjAy)
End Function
Function VbePjAyWh(A As Vbe, B As WhNm) As VBProject()
VbePjAyWh = AyMapPXInto(VbePjNyWh(A, B), "VbePj", A, VbePjAyWh)
End Function
Function VbePjNyWh(A As Vbe, B As WhNm) As String()
VbePjNyWh = AyWhNm(VbePjNy(A), B)
End Function
Function VbeSrcPth(A As Vbe)
Dim Pj As VBProject:
Set Pj = VbeFstQPj(A)
Dim Ffn$: Ffn = PjFfn(Pj)
If Ffn = "" Then Exit Function
VbeSrcPth = FfnPth(Pj.Filename)
End Function
Function VbeSrtRptFmt(A As Vbe) As String()
Dim Ay() As VBProject: Ay = VbePjAy(A)
Dim O$(), I, M As VBProject
For Each I In Ay
    Set M = I
    PushAy O, PjSrtRptFmt(M)
Next
VbeSrtRptFmt = O
End Function
Function VbePjFfn_Pj(A As Vbe, Ffn) As VBProject
Dim I
For Each I In A.VBProjects ' Cannot use VbePjAy(A), should use A.VBProjects
                           ' due to VbePjAy(X).FileName gives error
                           ' but (Pj in A.VBProjects).FileName is OK
    Debug.Print PjFfn(CvPj(I))
    If StrIsEq(PjFfn(CvPj(I)), Ffn) Then
        Set VbePjFfn_Pj = I
        Exit Function
    End If
Next
End Function
Function VbeWsFunNmzDupLines(A As Vbe) As Worksheet
Set VbeWsFunNmzDupLines = DrsWs(VbeDrsFunNmzDupLines(A))
End Function
Function VbeDrsFunNmzDupLines(A As Vbe) As Drs
'Nm AllLinesEq N Lines01....
Dim Drs As Drs
Set Drs = VbeFun12Drs(A)
Set Drs = DrsGpFlat(Drs, "Nm", "Lines")
Set Drs = DrsWhColGt(Drs, "N", 2)
Set Drs = DrsInsColBef(Drs, "N", "AllLinesEq")
Set Drs = VbeDrsFunNmzDupLines__1(Drs)
Set VbeDrsFunNmzDupLines = Drs
End Function
Private Function VbeDrsFunNmzDupLines__1(A As Drs) As Drs
'Update Col AllLinesEq
Dim O()
    Dim Dry(), J&, Dr
    Dry = A.Dry
    For Each Dr In Dry
        Dr(1) = VbeDrsFunNmzDupLines__2AllLinesIsEq(CvAy(Dr))
        Push O, Dr
    Next
Set VbeDrsFunNmzDupLines__1 = Drs(A.Fny, O)
End Function
Private Function VbeDrsFunNmzDupLines__2AllLinesIsEq(Dr()) As Boolean
'Nm AllLinesEq N Lines01....
'0  1          2 3
Dim L$, J%
L = Dr(3)
For J = 4 To UB(Dr)
    If Dr(J) <> L Then Exit Function
Next
VbeDrsFunNmzDupLines__2AllLinesIsEq = True
End Function
Sub VbeCompile(A As Vbe)
ItrDo A.VBProjects, "PjCompile"
End Sub
Sub VbeSav(A As Vbe)
ItrDo A.VBProjects, "PjSav"
End Sub
Sub VbeDmpIsSaved(A As Vbe)
Dim I As VBProject
For Each I In A.VBProjects
    Debug.Print I.Saved, I.BuildFileName
Next
End Sub
Sub VbeExport(A As Vbe)
Dim P As VBProject
For Each P In A.VBProjects
    PjExport P
Next
End Sub
Sub VbeSrcPthBrw(A As Vbe)
PthBrw VbeSrcPth(A)
End Sub
Sub VbeSrt(A As Vbe)
Dim I
For Each I In VbePjAy(A)
    PjSrt CvPj(I)
Next
End Sub
Sub VbeSrtRptBrw(A As Vbe)
Brw VbeSrtRptFmt(A)
End Sub
Function VbeMthPfx(A As Vbe) As String()

End Function
Function VbeSrc(A As Vbe) As String()
Dim P
For Each P In A.VBProjects
    PushAy VbeSrc, PjSrc(CvPj(P))
Next
End Function
Function VbeMth12Drs(A As Vbe) As Drs
Set VbeMth12Drs = Drs(Mth12DrFny, VbeMth12Dry(A))
End Function
Function VbeFun12Drs(A As Vbe) As Drs
Set VbeFun12Drs = Drs(Mth12DrFny, VbeFun12Dry(A))
End Function
Function VbeMth12Ws(A As Vbe) As Worksheet
Set VbeMth12Ws = DrsWs(VbeMth12Drs(A))
End Function
Function VbeMth12Dry(A As Vbe) As Variant()
VbeMth12Dry = AyMapFlat(VbePjAy(A), "PjMth12Dry")
End Function
Function VbeFun12Dry(A As Vbe) As Variant()
VbeFun12Dry = AyMapFlat(VbePjAy(A), "PjFun12Dry")
End Function
Function VbeMthDot(A As Vbe, Optional MthRe As RegExp, Optional MthExlAy$, Optional WhMdyAy, Optional WhMthKd0$, Optional PjRe As RegExp, Optional PjExlAy$, Optional MdRe As RegExp, Optional MdExlAy$)
Stop '
'Dim O$(), P
'For Each P In AyNz(VbePjAy(A, PjPatn, PjExlAy))
'    PushAy O, PjMthDot(CvPj(P), MthPatn, MthExlAy, MdPatn, MdExlAy, WhMdyA, WhMthKd0)
'Next
'VbeMthDot = O
End Function
Function VbeMnuBar(A As Vbe) As CommandBar
Set VbeMnuBar = A.CommandBars("Menu Bar")
End Function
Function VbeHasBar(A As Vbe, Nm$) As Boolean
VbeHasBar = ItrHasNm(A.CommandBars, Nm)
End Function
Sub VbeCrtBar(A As Vbe, Nm$)

End Sub
Function VbeBarNy(A As Vbe) As String()
VbeBarNy = ItrNy(A.CommandBars)
End Function
Sub VbeClsWin(A As Vbe, Optional ExcptWinTyAy)
Dim W As VBIDE.Window
If IsEmpty(ExcptWinTyAy) Then
    ItrDoSub A.Windows, "Close"
    Exit Sub
End If
End Sub
