Attribute VB_Name = "IdeMd"
Option Explicit
Sub MdEnsZ3DMthAsPrivate(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZ3DMthAsPrivate: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If IsPubZDashMthLin(L) Then
        Stop '
        'By = MthLin_EnsPrv(L)
        Debug.Print FmtQQ("MdEnsZ3DMthAsPrivate Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub
Private Sub MdEnsZZDashPrvMthAsPub(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If IsPrvZZDashMthLin(L) Then
        By = MthLin_EnsPub(L)
        Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) Lin(?) is change to Public: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub
Private Sub MdEnsZZDashPubMthAsPrv(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPubMthAsPrv Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If IsPubZZDashMthLin(L) Then
        Debug.Print L
        By = MthLin_EnsPrv(L)
        Debug.Print FmtQQ("MdEnsZZDashPubMthAsPrv: Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        'A.ReplaceLine J, By
    End If
Next
End Sub
Function MdMthWs(A As CodeModule) As Worksheet
Set MdMthWs = WsVis(SqWs(MdMthSq(A)))
End Function
Function MdMthDryWh(A As CodeModule, B As WhMth, Optional MthBrkOpt) As Variant()
MdMthDryWh = DryAddCC(SrcMthDryWh(MdBdyLy(A), B, C), MdPjNm(A), MdNm(A))
End Function
Sub Z_MdMthDry()
Brw DryFmtss(MdMthDry(CurMd))
End Sub
Function MdMthDry(A As CodeModule) As Variant()
MdMthDry = DryInsCC(SrcMthDry(MdBdyLy(A)), MdPjNm(A), MdNm(A))
End Function
Function MdMthKy(A As CodeModule, Optional IsWrap As Boolean) As String()
Dim PjN$: PjN = MdPjNm(A)
Dim MdN$: MdN = MdNm(A)
MdMthKy = SrcMthKy(MdSrc(A), PjN, MdN, IsWrap)
End Function
Function MdMthLinDry(A As CodeModule) As Variant()
MdMthLinDry = SrcMthLinDry(MdBdyLy(A))
End Function
Function MdMthLinDryWP(A As CodeModule) As Variant()
MdMthLinDryWP = SrcMthLinDryWP(MdBdyLy(A))
End Function
Sub MdSav(A As CodeModule)

End Sub
Function Md(MdDNm) As CodeModule
Dim A$: A = MdDNm
Dim P As VBProject
Dim MdNm$
    Dim L%
    L = InStr(A, ".")
    If L = 0 Then
        Set P = CurPj
        MdNm = A
    Else
        Dim PjNm$
        PjNm = Left(A, L - 1)
        Set P = Pj(PjNm)
        MdNm = Mid(A, L + 1)
    End If
Set Md = P.VBComponents(MdNm).CodeModule
End Function
Function MdMthLinAy(A As CodeModule) As String()
MdMthLinAy = SrcMthLinAy(MdSrc(A))
End Function
Function MdCmpTy(A As CodeModule) As vbext_ComponentType
MdCmpTy = A.Parent.Type
End Function
Function MdDNm$(A As CodeModule)
MdDNm = MdPjNm(A) & "." & MdNm(A)
End Function
Function MdDic(A As CodeModule, Optional ExlDcl As Boolean) As Dictionary
Set MdDic = SrcMthLinesDic(MdSrc(A), ExlDcl)
End Function
Function MdMthKeyLinesDic(A As CodeModule) As Dictionary
Set MdMthKeyLinesDic = SrcMthKeyLinesDic(MdSrc(A), MdPjNm(A), MdNm(A))
End Function
Function MdBdyLy(A As CodeModule) As String()
MdBdyLy = SplitCrLf(MdBdyLines(A))
End Function
Function MdHasNoLin(A As CodeModule) As Boolean
MdHasNoLin = A.CountOfLines = 0
End Function
Function MdBdyLines$(A As CodeModule)
If MdHasNoLin(A) Then Exit Function
MdBdyLines = A.Lines(A.CountOfDeclarationLines + 1, A.CountOfLines)
End Function
Function MdHasMth(A As CodeModule, MthNm) As Boolean
MdHasMth = SrcHasMth(MdBdyLy(A), MthNm)
End Function
Function MdHasTstSub(A As CodeModule) As Boolean
Dim I
For Each I In MdLy(A)
    If I = "Friend Sub Z__Tst()" Then MdHasTstSub = True: Exit Function
    If I = "Sub Z__Tst()" Then MdHasTstSub = True: Exit Function
Next
End Function
Function MdLines$(A As CodeModule)
With A
    If .CountOfLines = 0 Then Exit Function
    MdLines = .Lines(1, .CountOfLines)
End With
End Function
Function MdLy(A As CodeModule) As String()
MdLy = Split(MdLines(A), vbCrLf)
End Function
Function MdMthAy(A As CodeModule) As Mth()
Dim N
For Each N In MdMthNy(A)
    PushObj MdMthAy, Mth(A, N)
Next
End Function
Function MdMthAySel(A As CodeModule, B As WhMth)
End Function
Function MdMthAyWh(A As CodeModule, B As WhMth) As Mth()
Dim N
For Each N In AyNz(MdMthNyWh(A, B))
    PushObj MdMthAyWh, Mth(A, N)
Next
End Function
Function MdMthLno(A As CodeModule, MthNm) As Integer()
MdMthLno = AyAdd1(SrcMthNmIx(MdSrc(A), MthNm))
End Function
Function MdMthSq(A As CodeModule) As Variant()
MdMthSq = MthKy_Sq(MdMthKy(A, True))
End Function
Function MdMthFNyWhMth(A As CodeModule, B As WhMth) As String()
MdMthFNyWhMth = AyAddPfx(SrcMthNyWh(MdBdyLy(A), B), MdNm(A) & ".")
End Function
Function MdMthNyWh(A As CodeModule, B As WhMth) As String()
MdMthNyWh = AyAddPfx(SrcMthNyWh(MdBdyLy(A), B), MdNm(A) & ".")
End Function
Function MdMthFNyWh(A As CodeModule, B As WhMth) As String()
MdMthFNyWh = AyAddPfx(SrcMthFNyWh(MdBdyLy(A), B), MdNm(A) & ".")
End Function
Function MdMthNy(A As CodeModule) As String()
MdMthNy = AyAddPfx(SrcMthNy(MdBdyLy(A)), MdNm(A) & ".")
End Function
Function MdNm$(A As CodeModule)
MdNm = A.Parent.Name
End Function
Function MdPj(A As CodeModule) As VBProject
Set MdPj = A.Parent.Collection.Parent
End Function
Function MdPjNm$(A As CodeModule)
MdPjNm = MdPj(A).Name
End Function
Function MdRmk(A As CodeModule) As Boolean
Debug.Print "Rmk " & A.Parent.Name,
If IsMdAllRemarked(A) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To A.CountOfLines
    A.ReplaceLine J, "'" & A.Lines(J, 1)
Next
MdRmk = True
End Function
Function MdSrc(A As CodeModule) As String()
MdSrc = MdLy(A)
End Function
Function MdSrcExt$(A As CodeModule)
Dim O$
Select Case A.Parent.Type
Case vbext_ct_ClassModule: O = ".cls"
Case vbext_ct_Document: O = ".cls"
Case vbext_ct_StdModule: O = ".bas"
Case vbext_ct_MSForm: O = ".cls"
Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
End Select
MdSrcExt = O
End Function
Function MdSrcFfn$(A As CodeModule)
MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
End Function
Function MdSrcFn$(A As CodeModule)
MdSrcFn = MdNm(A) & MdSrcExt(A)
End Function
Function MdSrtRpt(A As CodeModule) As DCRslt
Dim P$, M$
M = MdNm(A)
P = MdPjNm(A)
MdSrtRpt = SrcSrtRpt(MdSrc(A), P, M)
End Function
Function MdSrtRptFmt(A As CodeModule) As String()
MdSrtRptFmt = SrcSrtRptFmt(MdSrc(A), MdPjNm(A), MdNm(A))
End Function
Function MdTyNm$(A As CodeModule)
MdTyNm = CmpTy_Nm(MdCmpTy(A))
End Function
Function MdUnRmk(A As CodeModule) As Boolean
Debug.Print "UnRmk " & A.Parent.Name,
If Not IsMdAllRemarked(A) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.ReplaceLine J, Mid(L, 2)
Next
MdUnRmk = True
End Function
Function MdXNm_Either(A) As Either
'Return ~.Left as MdDNm
'Or     ~.Right as PjNy() for those Pj holding giving Md
Dim P$, M$
Brk2Asg A, ".", P, M
If P <> "" Then
    MdXNm_Either = EitherL(A)
    Exit Function
End If
Dim Ny$()
Ny = VbePjNyWhMd(CurVbe, M)
If Sz(Ny) = 1 Then
    MdXNm_Either = EitherL(Ny(0) & "." & M)
    Exit Function
End If
MdXNm_Either = EitherR(Ny)
End Function
Function Md_FunNy_OfPfx_ZZDash(A As CodeModule) As String()
Dim J%, O$(), L$, L1$, Is_ZFun As Boolean
For J = 1 To A.CountOfLines
    Is_ZFun = True
    L = A.Lines(J, 1)
    Select Case True
    Case IsPfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case IsPfx(L, "Private Sub ZZ_")
        Is_ZFun = True
        L1 = RmvPfx(L, "Sub ")
    Case Else:
        Is_ZFun = False
    End Select

    If Is_ZFun Then
        Push O, TakNm(L1)
    End If
Next
Md_FunNy_OfPfx_ZZDash = O
End Function
Function MdFTLines$(A As CodeModule, X As FTNo)
Dim Cnt%: Cnt = FTNoLinCnt(X)
If Cnt = 0 Then Exit Function
MdFTLines = A.Lines(X.Fmno, Cnt)
End Function
Function MdFTLy(A As CodeModule, X As FTNo) As String()
MdFTLy = SplitCrLf(MdFTLines(A, X))
End Function
Function Md_TstSub_Lno%(A As CodeModule)
Dim J%
For J = 1 To A.CountOfLines
    If LinIsTstSub(A.Lines(J, 1)) Then Md_TstSub_Lno = J: Exit Function
Next
End Function
Function MdyIsSel(A$, MdySy$()) As Boolean
If Sz(MdySy) = 0 Then MdyIsSel = True: Exit Function
Dim Mdy
For Each Mdy In MdySy
    If Mdy = "Public" Then
        If A = "" Then MdyIsSel = True: Exit Function
    End If
    If A = Mdy Then MdyIsSel = True: Exit Function
Next
End Function
Function MdyShtMdy(A)
Dim O$
Select Case A
Case "", "Public":
Case "Private": O = "Prv"
Case "Friend": O = "Frd"
Case Else: Stop
End Select
MdyShtMdy = O
End Function
Function MdAyWhInTy(A() As CodeModule, WhInTyAy0$) As CodeModule()
Dim TyAy() As vbext_ComponentType, Md
TyAy = CvWhCmpTy(WhInTyAy0)
Dim O() As CodeModule
For Each Md In A
    If AyHas(TyAy, CvMd(Md).Parent.Type) Then PushObj O, Md
Next
MdAyWhInTy = O
End Function
Function MdTy(A As CodeModule) As vbext_ComponentType
MdTy = A.Parent.Type
End Function
Function MdAyWhMdy(A() As CodeModule, CmpTyAy0) As CodeModule()
'MdAyWhMdy = AyWhPredXP(A, "MdIsInCmpAy", CvCmpTyAy(CmpTyAy0))
End Function
Sub MdFmCntDlt(A As CodeModule, B() As FmCnt)
If Not IsFmCntInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub
Function MdyAy() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(0) = "Private"
    O(1) = "Friend"
    O(2) = "Public"
End If
MdyAy = O
End Function
Sub MdAddFun(A As CodeModule, Nm$, Lines)
MdAddIsFun A, Nm, Lines, IsFun:=True
End Sub
Sub MdAddSub(A As CodeModule, Nm$, Lines)
MdAddIsFun A, Nm, Lines, IsFun:=False
End Sub
Sub MdAddIsFun(A As CodeModule, Nm$, Lines, IsFun As Boolean)
Dim L$
    Dim B$
    B = IIf(IsFun, "Function", "Sub")
    L = FmtQQ("? ?()|?|End ?", B, Nm, Lines, B)
MdAppLines A, L
MthGo Mth(A, Nm)
End Sub
Sub MdAppLines(A As CodeModule, Lines$)
A.InsertLines A.CountOfLines + 1, Lines
End Sub
Sub MdClr(A As CodeModule, Optional IsSilent As Boolean)
With A
    If .CountOfLines = 0 Then Exit Sub
    If Not IsSilent Then Debug.Print FmtQQ("MdClr: Md(?) of lines(?) is cleared", MdNm(A), .CountOfLines)
    .DeleteLines 1, .CountOfLines
End With
End Sub
Sub MdCmp(A As CodeModule, B As CodeModule)
Dim A1 As Dictionary, B1 As Dictionary
    Set A1 = MdDic(A)
    Set B1 = MdDic(B)
Dim C As DCRslt
    C = DicCmp(A1, B1, MdDNm(A), MdDNm(B))
Brw DCRsltFmt(C)
End Sub
Sub MdCpy(A As CodeModule, ToPj As VBProject, Optional ShwMsg As Boolean)
Dim MdNm$
Dim FmPj As VBProject
    Set FmPj = MdPj(A)
    MdNm = A.Parent.Name
If PjHasCmp(ToPj, MdNm) Then
    Debug.Print FmtQQ("MdCpy: Md(?) exists in TarPj(?).  Skip copying", MdNm, ToPj.Name)
    Exit Sub
End If
Dim TmpFil$
    TmpFil = TmpFfn(".txt")
    Dim SrcCmp As VBComponent
    Set SrcCmp = A.Parent
    SrcCmp.Export TmpFil
    If SrcCmp.Type = vbext_ct_ClassModule Then
        FtRmvFst4Lines TmpFil
    End If
Dim TarCmp As VBComponent
    Set TarCmp = ToPj.VBComponents.Add(A.Parent.Type)
    TarCmp.CodeModule.AddFromFile TmpFil
Kill TmpFil
If ShwMsg Then Debug.Print FmtQQ("MdCpy: Md(?) is copied from SrcPj(?) to TarPj(?).", MdNm, FmPj.Name, ToPj.Name)
End Sub
Sub MdDlt(A As CodeModule)
Dim M$, P$, Pj As VBProject
    M = MdNm(A)
    Set Pj = MdPj(A)
    P = Pj.Name
Debug.Print FmtQQ("MdDlt: Before Md(?) is deleted from Pj(?)", M, P)
A.Parent.Collection.Remove A.Parent
'PjSav Pj
Debug.Print FmtQQ("MdDlt: After Md(?) is deleted from Pj(?)", M, P)
End Sub
Sub MdEndTrim(A As CodeModule, Optional ShwMsg As Boolean)
If A.CountOfLines = 0 Then Exit Sub
Dim N$: N = MdDNm(A)
Dim J%
While Trim(A.Lines(A.CountOfLines, 1)) = ""
    If ShwMsg Then Debug.Print FmtQQ("MdEndTrim: Lin(?) in Md(?) is removed due to it is blank", A.CountOfLines, N)
    A.DeleteLines A.CountOfLines, 1
    If A.CountOfLines = 0 Then Exit Sub
    If J > 1000 Then Stop
    J = J + 1
Wend
End Sub
Sub MdExport(A As CodeModule)
Dim F$: F = MdSrcFfn(A)
A.Parent.Export F
Debug.Print MdNm(A)
End Sub
Sub MdGo(A As CodeModule)
ClsWinExptImm
With A.CodePane
    .Show
    .Window.WindowState = vbext_ws_Maximize
End With
SendKeys "%WV"
End Sub
Sub MdGoMayLCC(Md As CodeModule, MayLCC As MayLCC)
MdGo Md
With MayLCC
    If .Som Then
        With .LCC
            Md.CodePane.TopLine = .Lno
            Md.CodePane.SetSelection .Lno, .C1, .Lno, .C2
        End With
    End If
End With
SendKeys "^{F4}"
End Sub
Sub MdRplCxt(A As CodeModule, Cxt$)
Dim N%: N = A.CountOfLines
MdClr A, IsSilent:=True
A.AddFromString Cxt
Debug.Print FmtQQ("MdRpl_Cxt: Md(?) of Ty(?) of Old-LinCxt(?) is replaced by New-Len(?) New-LinCnt(?).<-----------------", _
    MdDNm(A), MdTyNm(A), N, Len(Cxt), LinCnt(Cxt))
End Sub

Sub MdSrt(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
Debug.Print "Sorting: "; AlignL(Nm, 30); " ";
If MdNm(A) = "G_Tool" And MdPjNm(A) = "QTool" Then
    Debug.Print "<<<< Skipped"
    Exit Sub
End If
Dim NewLines$: NewLines = MdSrtedLines(A)
Dim Old$: Old = MdLines(A)
'Exit if same
    If Old = NewLines Then
        Debug.Print "<== Same"
        Exit Sub
    End If
Debug.Print "<-- Sorted";
'Delete
    Debug.Print FmtQQ("<--- Deleted (?) lines", A.CountOfLines);
    MdClr A, IsSilent:=True
'Add sorted lines
    A.AddFromString NewLines
    Md_Rmv_EmptyLines_AtEnd A
    Debug.Print "<----Sorted Lines added...."
End Sub
Sub Md_Gen_TstSub(A As CodeModule)
Md_Rmv_TstSub A
MdAppLines A, MdSubZLines(A)
End Sub
Sub Md_Mov_ToPj(A As CodeModule, ToPj As VBProject)
If MdNm(A) = "F__Tool" And CurPj.Name = "QTool" Then
    Debug.Print "Md(QTool.F__Tool) cannot be moved"
    Exit Sub
End If
MdCpy A, ToPj
MdDlt A
End Sub
Sub Md_Rmv_EmptyLines_AtEnd(A As CodeModule)
Dim J%
While A.CountOfLines > 1
    J = J + 1
    If J > 10000 Then Stop
    If Trim(A.Lines(A.CountOfLines, 1)) <> "" Then Exit Sub
    A.DeleteLines A.CountOfLines, 1
Wend
End Sub
Sub Md_Rmv_TstSub(A As CodeModule)
Dim L&, N&
L = Md_TstSub_Lno(A)
If L = 0 Then Exit Sub
Dim Fnd As Boolean, J%
For J = L + 1 To A.CountOfLines
    If IsPfx(A.Lines(J, 1), "End Sub") Then
        N = J - L + 1
        Fnd = True
        Exit For
    End If
Next
If Not Fnd Then Stop
A.DeleteLines L, N
End Sub
Sub MdRmvFC(A As CodeModule, B() As FmCnt)
If Not IsFmCntInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub
Private Sub Z_MdEndTrim()
Dim M As CodeModule: Set M = Md("ZZModule")
MdAppLines M, "  "
MdAppLines M, "  "
MdAppLines M, "  "
MdAppLines M, "  "
MdEndTrim M, ShwMsg:=True
Ass M.CountOfLines = 15
End Sub
Function MdMth9Dry(A As CodeModule) As Variant()
'Pj Md Mdy Ty Nm Sfx Prm Ret Rmk
'1  2  3   4  5  6   7   8   9
Dim Src$(), Dry()
Src = MdSrc(A)
Dry = SrcMth7Dry(Src)
Stop '
End Function
Function MdMth12Dry(A As CodeModule) As Variant()
Dim Pj$, Md$, Dry()
Pj = MdPjNm(A)
Md = MdNm(A)
Dry = SrcMth10Dry(MdSrc(A))
MdMth12Dry = DryInsCC(Dry, Pj, Md)
End Function
Function MdDryFun12(A As CodeModule) As Variant()
If Not IsStdMd(A) Then Exit Function
MdDryFun12 = MdMth12Dry(A)
End Function
Function MdSrtedLines$(A As CodeModule)
MdSrtedLines = SrcSrtedLines(MdSrc(A))
End Function
Function MdSubZLines$(A As CodeModule)
Dim Ny$(): Ny = Md_FunNy_OfPfx_ZZDash(A)
If Sz(Ny) = 0 Then Exit Function
Ny = AySrt(Ny)
Dim O$()
Dim Pfx$
If A.Parent.Type = vbext_ct_ClassModule Then
    Pfx = "Friend "
End If
Push O, ""
Push O, Pfx & "Sub Z__Tst()"
PushAy O, Ny
Push O, "End Sub"
MdSubZLines = Join(O, vbCrLf)
End Function
Sub MdMovMth(A As CodeModule, MthPatn$, ToMd As CodeModule)
Dim MthNy$(), M
Stop '
'MthNy = AyWhPatn(MdMthNy(A, "Pub"), MthPatn)
For Each M In AyNz(MthNy)
    MthMov Mth(A, M), ToMd
Next
End Sub
Function MdMthDot(A As CodeModule, Optional WhMdyAy, Optional WhKdAy) As String()
Stop '
'MdMthDot = SrcMthDot(MdBdyLy(A), WhMdyA, WhKdAy)
End Function
Function MdMthLinesWithRmk$(A As CodeModule, MthNm)
MdMthLinesWithRmk = SrcMthLinesWithRmk(MdBdyLy(A), MthNm)
End Function
Function MdMthLines$(A As CodeModule, M$)
MdMthLines = MthLines(Mth(A, M))
End Function
Private Sub ZZ_MdMthDic()
Dim D As Dictionary
Set D = MdMthDic(CurMd)
Stop
End Sub
Function MdMthLinesDic(A As CodeModule) As Dictionary
Set MdMthLinesDic = SrcMthLinesDic(MdSrc(A))
End Function
Function MdMthDic(A As CodeModule) As Dictionary
Set MdMthDic = SrcMthDic(MdBdyLy(A), MdPjNm(A), MdNm(A))
End Function
Private Sub Z_MdMthDic()
DicBrw MdMthDic(CurMd)
End Sub
Sub MdClsWin(A As CodeModule)
A.CodePane.Window.Close
End Sub
Function MdMthPfx(A As CodeModule) As String()

End Function
