Attribute VB_Name = "M_Tool3"
'Option Explicit
'Private Type Either
'    IsLeft As Boolean
'    Left As Variant
'    Right As Variant
'End Type
'Private Type FmToLno
'    FmLno As Integer
'    ToLno As Integer
'End Type
'Private Type DCRslt
'    Nm1 As String
'    Nm2 As String
'    AExcess As New Dictionary
'    BExcess As New Dictionary
'    ADif As New Dictionary
'    BDif As New Dictionary
'    Sam As New Dictionary
'End Type
'Private Type DicPair
'    A As Dictionary
'    B  As Dictionary
'End Type
'Private Type S1S2
'    S1 As String
'    S2 As String
'End Type
'Private Type SyPair
'    Sy1() As String
'    Sy2() As String
'End Type
'Private Type MdSrtRpt
'    MdNy() As String
'    RptDic As Dictionary ' K is Module Name, V is DicCmpRsltLy
'End Type
'Private Type LCC
'    Lno As Integer
'    C1 As Integer
'    C2 As Integer
'End Type
'Private Type LCCOpt
'    Som As Boolean
'    LCC As LCC
'End Type
'Enum eCmpTySelection
'    eMdOnly = 1
'    eClsOnly = 2
'    eBothMdAndCls = 3
'End Enum
'Private Fso As New Scripting.FileSystemObject
'
'Sub AAA()
'Srt_M_Tool
'End Sub
'
'Sub Add_Cls(Nm$)
'ZPj_Add_Mbr ZCurPj, Nm, vbext_ct_ClassModule
'End Sub
'
'Sub Add_CurVbe_QToolRf()
'Dim I, P As VBProject
'For Each I In ZCurVbe_PjAy
'    Set P = I
'    ZPj_AddRf P, "QTool"
'Next
'End Sub
'
'Sub Add_Fun(FunNm$)
'ZAdd_Fun_or_Sub FunNm, IsFun:=True
'End Sub
'
'Sub Add_Md(Nm$)
'ZPj_Add_Mbr ZCurPj, Nm, vbext_ct_StdModule
'End Sub
'
'Sub Add_Sub(SubNm$)
'ZAdd_Fun_or_Sub SubNm, IsFun:=False
'End Sub
'
'Sub Brw_Pj_SrcPth()
'ZPj_SrcPthBrw ZCurPj
'End Sub
'
'Sub Brw_Pj_SrtRpt()
'ZAyBrw ZPj_SrtRptLy(ZCurPj)
'End Sub
'
'Sub Brw_Vbe_SrtRpt()
'ZAyBrw ZVbe_SrtRptLy(ZCurVbe)
'End Sub
'
'Sub Brw_Md_MthKy()
'ZAyBrw ZMd_MthKy(ZCurMd, IsSngLinFmt:=True)
'End Sub
'Sub Brw_Pj_MthKy()
'ZAyBrw ZPj_MthKy(ZCurPj, IsSngLinFmt:=True)
'End Sub
'Sub Brw_Vbe_MthKy()
'ZAyBrw ZVbe_MthKy(ZCurVbe, IsSngLinFmt:=True)
'End Sub
'
'Sub Brw_Pj_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
'ZAyBrw ZCurPj_MthNy(MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
'End Sub
'
'Sub Cls_Win()
'Dim W As VBIDE.Window
'For Each W In ZCurVbe.Windows
'    W.Close
'Next
'End Sub
'
'Sub Compile_Pj()
'ZPj_Compile ZCurPj
'End Sub
'
'Sub Compile_Vbe()
'ZAyDo ZCurVbe_PjAy, "ZPj_Compile"
'End Sub
'
'Sub Dlt_Md()
'If MsgBox(ZFmtQQ("Delete this Md[?]", ZCurMdNm), vbYesNo + vbDefaultButton2) <> vbYes Then Exit Sub
'ZCurPj.VBComponents.Remove ZCurCmp
'End Sub
'
'Sub Export_Pj()
'ZPj_Export ZCurPj
'End Sub
'
''=============================
'Sub Export_Vbe()
'ZVbe_Export ZCurVbe
'End Sub
'
'Sub Gen_Md_TstSub()
'ZMd_Gen_TstSub ZCurMd
'End Sub
'
'Sub Gen_Pj_ConstructorMd()
'Stop '
'End Sub
'
'Sub Gen_Pj_TstClass()
'ZPj_Gen_TstClass ZCurPj
'End Sub
'
'Sub Gen_Pj_TstSub()
'ZPj_Gen_TstSub ZCurPj
'End Sub
'
'Sub Gen_Vbe_TstClass()
'End Sub
'Private Property Get ZEitherR(A) As Either
'ZAsg A, ZEitherR.Right
'End Property
'Private Property Get ZEitherL(A) As Either
'ZAsg A, ZEitherL.Left
'ZEitherL.IsLeft = True
'End Property
'Private Property Get ZPjMbrDotNm_Either(A) As Either
''Return ~.Left as PjMbrDotNm
''Or     ~.Right as PjNy() for those Pj holding giving Md
'Dim P$, M$
'ZBrk2_Asg A, ".", P, M
'If P <> "" Then
'    ZPjMbrDotNm_Either = ZEitherL(A)
'End If
'Dim Ny$()
'Ny = ZCurVbe_MdPjNy(M)
'If ZSz(Ny) = 1 Then
'    ZPjMbrDotNm_Either = ZEitherL(Ny(0))
'    Exit Property
'End If
'ZPjMbrDotNm_Either = ZEitherR(Ny)
'End Property
'Sub Go_Mbr(PjMbrDotNm$)
'Dim E As Either
'E = ZPjMbrDotNm_Either(PjMbrDotNm)
'If E.IsLeft Then
'    ZMd_Go ZMd(E.Left)
'    Exit Sub
'End If
'Dim Ny$()
'    Ny = E.Right
'If ZSz(Ny) = 0 Then
'    Debug.Print PjMbrDotNm; "<-- No such module"
'    Stop '
'    Exit Sub
'End If
'Dim I
'For Each I In Ny
'    Debug.Print "Go_Mbr """; I; "."; PjMbrDotNm
'Next
'End Sub
'
'Sub Go_Mth(PjMdMthDotNm$)
'Dim Md As CodeModule
'Dim MthNm$
'ZPjMdMthDotNm_BrkAsg PjMdMthDotNm, _
'    Md, MthNm
'Dim L As LCCOpt
'    L = ZMdMth_LCCOpt(Md, MthNm)
'ZMd_GoLCCOpt Md, L
'End Sub
'
'Sub Go_Pj(PjNm$)
'ZPj_Go ZPj(PjNm)
'End Sub
'
'Sub Lis_Md()
'Dim A$()
'    A = ZCurPj_MbrNy
'    A = ZAySrt(A)
'    A = ZAyAddPfx(A, "Go_Mbr """)
'ZAyDmp A
'End Sub
'
'Sub Lis_Md_Mth(Optional MthNmPatn$ = ".")
'ZAyDmp ZAyAddPfx(ZCurMd_MthNy(MthNmPatn), ZCurMdNm & ".")
'End Sub
'
'Sub Lis_Mth(Optional MthNmPatn$ = ".")
'Lis_Vbe_Mth MthNmPatn
'End Sub
'
'Sub Lis_Pj()
'Dim A$()
'    A = ZCurVbe_PjNy
'    A = ZAyAddPfx(A, "Go_Pj """)
'ZAyDmp A
'End Sub
'
'Sub Lis_Pj_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
'Dim A$()
'    A = ZCurPj_MthNy(MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
'    A = ZAySrt(A)
'    A = ZAyAddPfx(A, "Go_Mth """)
'ZAyDmp A
'End Sub
'
'Sub Lis_Vbe_Mth(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".")
'Dim A$()
'    A = ZCurVbe_MthNy(MthNmPatn:=MthNmPatn, MdNmPatn:=MdNmPatn)
'    A = ZAySrt(A)
'ZAyDmp A
'End Sub
'
'Sub Mov_Md_ToPj(ToPjNm$)
'If ZCurPjNm = ToPjNm Then
'    Debug.Print ZFmtQQ("Mov_Md: ToPjNm(?) cannot be CurPjNm", ToPjNm)
'    Exit Sub
'End If
'ZMd_Mov_ToPj ZCurMd, ZPj(ToPjNm)
'End Sub
'
'Private Sub ZMd_Mov_ToPj(A As CodeModule, ToPj As VBProject)
'ZMd_Cpy_ToPj A, ToPj
'ZMd_Dlt A
'End Sub
'
'Sub Mov_MdLik_ToPj(MdLikNm$, ToPjNm$)
'Dim Ay() As CodeModule: Ay = ZCurPj_MbrAyLik(MdLikNm)
'If ZSz(Ay) = 0 Then Exit Sub
'Dim I, P As VBProject
'Set P = ZPj(ToPjNm)
'For Each I In Ay
'    ZMd_Mov_ToPj ZCvMd(I), P
'Next
'ZClsWinExcept_Module_A_1
'End Sub
'
'Sub Ren_Md(NewNm$)
'If ZCurPj_HasCmp(NewNm) Then
'    MsgBox ZFmtQQ("Md(?) exists in CurPj(?).  Cannot rename.", NewNm, ZCurPjNm), , "M_A:RenMd"
'    Exit Sub
'End If
'ZCurMd.Name = NewNm
'End Sub
'
'Sub Rmk_All()
'Dim I, Md As CodeModule
'Dim NRmk%, Skip%
'For Each I In ZMbrAy
'    Set Md = I
'    If ZMd_Rmk(Md) Then
'        NRmk = NRmk + 1
'    Else
'        Skip = Skip + 1
'    End If
'Next
'Debug.Print "NRmk"; NRmk
'Debug.Print "SKip"; Skip
'End Sub
'
'Function Shw_CurPj_SrtRptWb(Optional Vis As Boolean) As Workbook
'ZPj_SrtRptWb ZCurPj, Vis
'End Function
'
'Sub Srt_M_Tool()
'Dim P As VBProject
'Dim Md As CodeModule
'Dim Src$()
'Dim Cxt$
'Set P = ZPj("QTool")
'Set Md = ZPj_Md(P, "M_Tool")
'Src = ZMd_Src(Md)
'Cxt = ZSrc_SrtedLines(Src)
'ZPj_Ens_Md P, "M_Tool1", Cxt
'End Sub
'
'Sub Srt_Md()
'ZMd_Srt ZCurMd
'End Sub
'
'Sub Srt_Pj()
'ZPj_Srt ZCurPj
'End Sub
'
'Sub UnRmk_All()
'Dim I, Md As CodeModule
'Dim NUnRmk%, Skip%
'For Each I In ZMbrAy
'    Set Md = I
'    If ZMd_UnRmk(Md) Then
'        NUnRmk = NUnRmk + 1
'    Else
'        Skip = Skip + 1
'    End If
'Next
'Debug.Print "NUnRmk"; NUnRmk
'Debug.Print "SKip"; Skip
'End Sub
'
'Private Property Get ZAySrt__Ix&(Ay, V, Des As Boolean)
'Dim I, O&
'If Des Then
'    For Each I In Ay
'        If V > I Then ZAySrt__Ix = O: Exit Property
'        O = O + 1
'    Next
'    ZAySrt__Ix = O
'    Exit Property
'End If
'For Each I In Ay
'    If V < I Then ZAySrt__Ix = O: Exit Property
'    O = O + 1
'Next
'ZAySrt__Ix = O
'End Property
'
'Private Property Get ZDCRslt_Ly__AExcess(A As Dictionary) As S1S2()
'If A.Count = 0 Then Exit Function
'Dim O() As S1S2, K
'For Each K In A.Keys
'    ZS1S2_Push O, ZS1S2(K & vbCrLf & ZLines_UnderLin(K) & vbCrLf & A(K), "!" & "Er AExcess")
'Next
'ZDCRslt_Ly__AExcess = O
'End Property
'
'Private Property Get ZDCRslt_Ly__BExcess(A As Dictionary) As S1S2()
'If A.Count = 0 Then Exit Function
'Dim O() As S1S2, K
'For Each K In A.Keys
'    ZS1S2_Push O, ZS1S2("!" & "Er BExcess", K & vbCrLf & ZLines_UnderLin(K) & vbCrLf & A(K))
'Next
'ZDCRslt_Ly__BExcess = O
'End Property
'
'Private Property Get ZDCRslt_Ly__Dif(A As Dictionary, B As Dictionary) As S1S2()
'If A.Count <> B.Count Then Stop
'If A.Count = 0 Then Exit Function
'Dim O() As S1S2, K, S1$, S2$
'For Each K In A
'    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ZLines_UnderLin(K) & vbCrLf & A(K)
'    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & ZLines_UnderLin(K) & vbCrLf & B(K)
'    ZS1S2_Push O, ZS1S2(S1, S2)
'Next
'ZDCRslt_Ly__Dif = O
'End Property
'
'Private Property Get ZDCRslt_Ly__Sam(A As Dictionary) As S1S2()
'If A.Count = 0 Then Exit Function
'Dim O() As S1S2, K
'For Each K In A.Keys
'    ZS1S2_Push O, ZS1S2("*Same", K & vbCrLf & ZLines_UnderLin(K) & vbCrLf & A(K))
'Next
'ZDCRslt_Ly__Sam = O
'End Property
'
'Sub ZZ__Tst()
'ZZ_Dcl_BefAndAft_Srt
'ZZ_PjSrtRptWb
'ZZ_Shw_CurPj_SrtRptWb
'ZZ_ZCurMdNm
'ZZ_ZCurVbe_PjNy
'ZZ_ZMd_Gen_TstSub
'ZZ_ZMd_Rmv_TstSub
'ZZ_ZMd_SrtedLines
'ZZ_ZMd_TstSub_BdyLines
'ZZ_ZMd_TstSub_Lno
'ZZ_ZPj
'ZZ_ZPj_MthS1S2Ay
'ZZ_ZPj_SrtRptLy
'ZZ_ZPj_TstClass_Bdy
'ZZ_ZS1S2Ay_FmtLy
'ZZ_ZSrc_DclLinCnt
'ZZ_ZSrc_DclLines
'ZZ_ZSrc_MthS1S2Ay
'ZZ_ZSrc_SrtRptLy
'ZZ_ZSrc_SrtedBdyLines
'ZZ_ZSrc_SrtedLines
'ZZ_ZSrc_SrtedLy
'End Sub
'
'Private Sub ZBrk2_Asg(A, Sep$, O1$, O2$)
'Dim P%: P = InStr(A, Sep)
'If P = 0 Then
'    O1 = ""
'    O2 = Trim(A)
'Else
'    O1 = Trim(Left(A, P - 1))
'    O2 = Trim(Mid(A, P + 1))
'End If
'End Sub
'
'Private Sub ZAdd_Fun_or_Sub(Nm$, IsFun As Boolean)
'Dim L$
'    Dim A$
'    A = IIf(IsFun, "Function", "Sub")
'    L = ZFmtQQ("? ?()|End ?", A, Nm, A)
'With ZMd(Nm)
'    .InsertLines .CountOfLines + 1, L
'End With
'Go_Mth Nm
'End Sub
'
'Private Property Get ZAlignL$(A, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
'Const CSub$ = "ZAlignL"
'If ErIfNotEnoughWdt And DoNotCut Then
'    Stop
'    'Er CSub, "Both {ErIfNotEnoughWdt} and {DontCut} cannot be True", ErIfNotEnoughWdt, DoNotCut
'End If
'Dim S$: S = ZToStr(A)
'ZAlignL = ZStrAlignL(S, W, ErIfNotEnoughWdt, DoNotCut)
'End Property
'
'Private Sub ZAsg(V, OV)
'If IsObject(V) Then
'   Set OV = V
'Else
'   OV = V
'End If
'End Sub
'
'Private Sub ZAss(A As Boolean)
'If Not A Then Stop
'End Sub
'
'Private Property Get ZAyAddPfx(Ay, Pfx) As String()
'If ZSz(Ay) = 0 Then Exit Property
'Dim O$(), I
'For Each I In Ay
'    ZPush O, Pfx & I
'Next
'ZAyAddPfx = O
'End Property
'
'Private Property Get ZAyAlignL(Ay) As String()
'Dim W%: W = ZAyWdt(Ay) + 1
'If ZSz(Ay) = 0 Then Exit Property
'Dim O$(), I
'For Each I In Ay
'    ZPush O, ZAlignL(I, W)
'Next
'ZAyAlignL = O
'End Property
'
'Private Sub ZAyBrw(Ay)
'ZStr_Brw Join(Ay, vbCrLf)
'End Sub
'
'Private Sub ZAyDmp(Ay)
'If ZSz(Ay) = 0 Then Exit Sub
'Dim I
'For Each I In Ay
'    Debug.Print I
'Next
'End Sub
'
'Private Sub ZAyDo(Ay, DoMthNm$)
'If ZSz(Ay) = 0 Then Exit Sub
'Dim I
'For Each I In Ay
'    Run DoMthNm, I
'Next
'End Sub
'
'Private Property Get ZAyFstNEle(Ay, N&)
'Dim O: O = Ay
'ReDim Preserve O(N - 1)
'ZAyFstNEle = O
'End Property
'
'Private Property Get ZAyHas(Ay, Itm) As Boolean
'Dim I: If ZSz(Ay) = 0 Then Exit Property
'For Each I In Ay
'    If I = Itm Then ZAyHas = True: Exit Property
'Next
'End Property
'
'Private Property Get ZAyIns(Ay, Optional Ele, Optional At&)
'Const CSub$ = "AyIns"
'Dim N&: N = ZSz(Ay)
'If 0 > At Or At > N Then
'    Stop
'End If
'Dim O
'    O = Ay
'    ReDim Preserve O(N)
'    Dim J&
'    For J = N To At + 1 Step -1
'        ZAsg O(J - 1), O(J)
'    Next
'    O(At) = Ele
'ZAyIns = O
'End Property
'
'Private Property Get ZAyLasEle(Ay)
'ZAyLasEle = Ay(ZUB(Ay))
'End Property
'
'Private Property Get ZAyMinus(A, B)
'If ZSz(B) = 0 Or ZSz(A) = 0 Then ZAyMinus = A: Exit Property
'Dim O: O = A: Erase O
'Dim B1: B1 = B
'Dim V
'For Each V In A
'    If ZAyHas(B1, V) Then
'        B1 = ZAyRmvEle(B1, V)
'    Else
'        ZPush O, V
'    End If
'Next
'ZAyMinus = O
'End Property
'
'Private Property Get ZAyMinusAp(Ay, ParamArray AyAp())
'Dim O
'If ZSz(Ay) = 0 Then O = Ay: Erase O: GoTo X
'O = Ay
'Dim Av(): Av = AyAp
'Dim Ay1, V
'For Each Ay1 In Av
'    O = ZAyMinus(O, Ay1)
'    If ZSz(O) = 0 Then GoTo X
'Next
'X:
'ZAyMinusAp = O
'End Property
'
'Private Property Get ZAyPair_Dic(A1, A2) As Dictionary
'Dim N1&, N2&
'N1 = ZSz(A1)
'N2 = ZSz(A2)
'If N1 <> N2 Then Stop
'Dim O As New Dictionary
'Dim J&
'If ZSz(A1) = 0 Then GoTo X
'For J = 0 To N1 - 1
'    O.Add A1(J), A2(J)
'Next
'X:
'Set ZAyPair_Dic = O
'End Property
'
'Private Property Get ZAyRmvEle(Ay, M)
'Dim O, V: O = Ay: Erase O
'For Each V In Ay
'    If V <> M Then ZPush O, M
'Next
'ZAyRmvEle = O
'End Property
'
'Private Property Get ZAyRmvEmp(Ay)
'If ZSz(Ay) = 0 Then ZAyRmvEmp = Ay: Exit Property
'Dim O: O = Ay: Erase O
'Dim I
'For Each I In Ay
'    If Not ZIs_Emp(I) Then ZPush O, I
'Next
'ZAyRmvEmp = O
'End Property
'
'Private Property Get ZAySqV(Ay) As Variant()
'If ZSz(Ay) = 0 Then Exit Property
'Dim O(), R&
'ReDim O(1 To ZSz(Ay), 1 To 1)
'R = 0
'Dim V
'For Each V In Ay
'    R = R + 1
'    O(R, 1) = V
'Next
'ZAySqV = O
'End Property
'
'Private Property Get ZAySrt(Ay, Optional Des As Boolean)
'If ZSz(Ay) = 0 Then ZAySrt = Ay: Exit Property
'Dim Ix&, V, J&
'Dim O: O = Ay: Erase O
'ZPush O, Ay(0)
'For J = 1 To ZUB(Ay)
'    O = ZAyIns(O, Ay(J), ZAySrt__Ix(O, Ay(J), Des))
'Next
'ZAySrt = O
'End Property
'
'Private Property Get ZAySrtInToIxAy_Ix&(Ix&(), A, V, Des As Boolean)
'Dim I, O&
'If Des Then
'    For Each I In Ix
'        If V > A(I) Then ZAySrtInToIxAy_Ix& = O: Exit Property
'        O = O + 1
'    Next
'    ZAySrtInToIxAy_Ix& = O
'    Exit Property
'End If
'For Each I In Ix
'    If V < A(I) Then ZAySrtInToIxAy_Ix& = O: Exit Property
'    O = O + 1
'Next
'ZAySrtInToIxAy_Ix& = O
'End Property
'
'Private Property Get ZAySrtIntoIxAy(Ay, Optional Des As Boolean) As Long()
'If ZSz(Ay) = 0 Then Exit Property
'Dim Ix&, V, J&
'Dim O&():
'ZPush O, 0
'For J = 1 To ZUB(Ay)
'    O = ZAyIns(O, J, ZAySrtInToIxAy_Ix(O, Ay, Ay(J), Des))
'Next
'ZAySrtIntoIxAy = O
'End Property
'
'Private Property Get ZAyUniqAy(Ay)
'Dim O: O = Ay: Erase O
'If ZSz(Ay) > 0 Then
'    Dim I
'    For Each I In Ay
'        ZPushNoDup O, I
'    Next
'End If
'ZAyUniqAy = O
'End Property
'
'Private Property Get ZAyWdt%(Ay)
'Dim W%, I: If ZSz(Ay) = 0 Then Exit Property
'For Each I In Ay
'    W = ZMax(Len(I), W)
'Next
'ZAyWdt = W
'End Property
'
'Private Property Get ZAyWhFmTo(Ay, FmIx, ToIx)
'Dim O: O = Ay: Erase O
'Dim J&
'For J = FmIx To ToIx
'    ZPush O, Ay(J)
'Next
'ZAyWhFmTo = O
'End Property
'
'Private Sub ZAyWrt(Ay, Ft$)
'ZStr_Wrt ZJnCrLf(Ay), Ft
'End Sub
'
'Private Sub ZClsWinExcept_Module_A_1()
'Dim W As VBIDE.Window
'For Each W In ZCurVbe.Windows
'    If W.Type = vbext_wt_CodeWindow Then
'        If W.Caption <> "Lib_XXX.xlam - A_1 (Code)" Then
'            W.Close
'        End If
'    End If
'Next
'End Sub
'
'Private Property Get ZCmpTy_Nm$(A As vbext_ComponentType)
'Dim O$
'Select Case A
'Case vbext_ct_ClassModule: O = "*Cls"
'Case vbext_ct_StdModule: O = "*Md"
'Case Else: Stop
'End Select
'ZCmpTy_Nm = O
'End Property
'
'Private Sub ZCmp_Rmv(A As VBComponent)
'A.Collection.Remove A
'End Sub
'
'Private Property Get ZCurCmp() As VBComponent
'Set ZCurCmp = ZCurMd.Parent
'End Property
'
'Private Property Get ZCurMd() As CodeModule
'Set ZCurMd = ZCurVbe.ActiveCodePane.CodeModule
'End Property
'
'Private Property Get ZCurMdNm$()
'ZCurMdNm = ZCurCmp.Name
'End Property
'
'Private Property Get ZCurMd_MthNy(Optional MthNmPatn$ = ".") As String()
'ZCurMd_MthNy = ZMd_MthNy(ZCurMd, MthNmPatn)
'End Property
'
'Private Property Get ZCurPj() As VBProject
'Set ZCurPj = ZCurVbe.ActiveVBProject
'End Property
'
'Private Property Get ZCurPjNm$()
'ZCurPjNm = ZCurPj.Name
'End Property
'
'Private Property Get ZCurPj_Cmp(Nm) As VBComponent
'Set ZCurPj_Cmp = ZPj_Cmp(ZCurPj, Nm)
'End Property
'
'Private Property Get ZCurPj_HasCmp(Nm$) As Boolean
'ZCurPj_HasCmp = ZPj_HasCmp(ZCurPj, Nm)
'End Property
'
'Private Property Get ZCurPj_MbrNy() As String()
'ZCurPj_MbrNy = ZPj_MbrNyLik(ZCurPj, "*")
'End Property
'
'Private Property Get ZCurPj_MbrNyLik(MdLikNm$) As String()
'ZCurPj_MbrNyLik = ZPj_MbrNyLik(ZCurPj, MdLikNm)
'End Property
'
'Private Property Get ZCurPj_MbrAyLik(MdLikNm$) As CodeModule()
'ZCurPj_MbrAyLik = ZPj_MbrAyLik(ZCurPj, MdLikNm)
'End Property
'
'Private Property Get ZCurPj_MthNy(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".") As String()
'ZCurPj_MthNy = ZPj_MthNy(ZCurPj, MthNmPatn, MdNmPatn)
'End Property
'
'Private Property Get ZCurVbe() As VBE
'Set ZCurVbe = Excel.Application.VBE
'End Property
'
'Private Property Get ZCurVbe_MthNy(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".") As String()
'ZCurVbe_MthNy = ZVbe_MthNy(ZCurVbe, MthNmPatn, MdNmPatn)
'End Property
'
'Private Property Get ZCurVbe_PjAy() As VBProject()
'ZCurVbe_PjAy = ZVbe_PjAy(ZCurVbe)
'End Property
'
'Private Property Get ZCurVbe_PjNy() As String()
'ZCurVbe_PjNy = ZVbe_PjNy(ZCurVbe)
'End Property
'
'Private Property Get ZCvMd(A) As CodeModule
'Set ZCvMd = A
'End Property
'
'Private Property Get ZCvPj(I) As VBProject
'Set ZCvPj = I
'End Property
'
'Private Property Get ZDCRslt_IsSam(A As DCRslt) As Boolean
'With A
'If .ADif.Count > 0 Then Exit Property
'If .BDif.Count > 0 Then Exit Property
'If .AExcess.Count > 0 Then Exit Property
'If .BExcess.Count > 0 Then Exit Property
'End With
'ZDCRslt_IsSam = True
'End Property
'
'Private Property Get ZDCRslt_Ly(A As DCRslt, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As String()
'With A
'Dim A1() As S1S2: A1 = ZDCRslt_Ly__AExcess(.AExcess)
'Dim A2() As S1S2: A2 = ZDCRslt_Ly__BExcess(.BExcess)
'Dim A3() As S1S2: A3 = ZDCRslt_Ly__Dif(.ADif, .BDif)
'Dim A4() As S1S2: A4 = ZDCRslt_Ly__Sam(.Sam)
'End With
'Dim O() As S1S2
'ZS1S2_Push O, ZS1S2(Nm1, Nm2)
'O = ZS1S2Ay_Add(O, A1)
'O = ZS1S2Ay_Add(O, A2)
'O = ZS1S2Ay_Add(O, A3)
'O = ZS1S2Ay_Add(O, A4)
'ZDCRslt_Ly = ZS1S2Ay_FmtLy(O)
'End Property
'
'Private Property Get ZDftMdByMdNm(MdNm$) As CodeModule
'If MdNm = "" Then
'    Set ZDftMdByMdNm = ZCurMd
'Else
'    Set ZDftMdByMdNm = ZMd(MdNm)
'End If
'End Property
'
'Private Property Get ZDicPair_SamKeyDifValPair(A As Dictionary, B As Dictionary) As DicPair
'Dim K, A1 As New Dictionary, B1 As New Dictionary
'For Each K In A.Keys
'    If B.Exists(K) Then
'        If A(K) <> B(K) Then
'            A1.Add K, A(K)
'            B1.Add K, B(K)
'        End If
'    End If
'Next
'With ZDicPair_SamKeyDifValPair
'    Set .A = A1
'    Set .B = B1
'End With
'End Property
'
'Private Property Get ZDic_Clone(A As Dictionary) As Dictionary
'Dim O As New Dictionary, K
'For Each K In A.Keys
'    O.Add K, A(K)
'Next
'Set ZDic_Clone = O
'End Property
'
'Private Property Get ZDic_Cmp(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As DCRslt
'Dim O As DCRslt
'Set O.AExcess = ZDic_Minus(A, B)
'Set O.BExcess = ZDic_Minus(B, A)
'Set O.Sam = ZDic_Sam(A, B)
'With ZDicPair_SamKeyDifValPair(A, B)
'    Set O.ADif = .A
'    Set O.BDif = .B
'End With
'O.Nm1 = Nm1
'O.Nm2 = Nm2
'ZDic_Cmp = O
'End Property
'
'Private Property Get ZDic_Minus(A As Dictionary, B As Dictionary) As Dictionary
'If A.Count = 0 Then Set ZDic_Minus = New Dictionary: Exit Property
'If B.Count = 0 Then Set ZDic_Minus = ZDic_Clone(A): Exit Property
'Dim O As New Dictionary, K
'For Each K In A.Keys
'   If Not B.Exists(K) Then O.Add K, A(K)
'Next
'Set ZDic_Minus = O
'End Property
'
'Private Property Get ZDic_Sam(A As Dictionary, B As Dictionary) As Dictionary
'Dim O As New Dictionary
'If A.Count = 0 Or B.Count = 0 Then GoTo X
'Dim K
'For Each K In A.Keys
'    If B.Exists(K) Then
'        If A(K) = B(K) Then
'            O.Add K, A(K)
'        End If
'    End If
'Next
'X: Set ZDic_Sam = O
'End Property
'
'Private Property Get ZDic_Wb(A As Dictionary, Optional Vis As Boolean) As Workbook
''Assume each dic keys is name and each value is lines
''Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
'ZAss ZIs_Dic_AllKeyIsNm(A)
'ZAss ZIs_Dic_AllValIsStr(A)
'Dim K, ThereIsSheet1 As Boolean
'Dim O As Workbook: Set O = ZNewWb
'Dim Ws As Worksheet
'For Each K In A.Keys
'    If K = "Sheet1" Then
'        Set Ws = O.Sheets("Sheet1")
'        ThereIsSheet1 = True
'    Else
'        Set Ws = O.Sheets.Add
'        Ws.Name = K
'    End If
'    Ws.Range("A1").Value = ZLines_SqV(A(K))
'Next
'X: Set Ws = O
'If Vis Then O.Application.Visible = True
'End Property
'
'Private Sub ZDotDotNm_BrkAsg(A, O1$, O2$, O3$)
'Dim Ay$(): Ay = Split(A, ".")
'Select Case ZSz(Ay)
'Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
'Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
'Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
'Case Else: Stop
'End Select
'End Sub
'
'Private Property Get ZEmp_RfAy() As Reference()
'End Property
'
''Function DftFfn(Ffn0, Optional Ext$ = ".txt", Optional Pth0$, Optional Fdr$)
''If Ffn0 <> "" Then DftFfn = Ffn0: Exit Function
''Dim Pth$: Pth = DftPth(Pth0)
''DftFfn = Pth & TmpNm & Ext
''End Function
''Function DftPth$(Optional Pth0$, Optional Fdr$)
''If Pth0 <> "" Then DftPth = Pth0: Exit Function
''DftPth = TmpPth(Fdr)
''End Function
''Function FfnAddFnSfx(A$, Sfx$)
''FfnAddFnSfx = FfnRmvExt(A) & Sfx & FfnExt(A)
''End Function
'Private Sub ZFfn_CpyToPth(A, ToPth$, Optional OvrWrt As Boolean)
'Fso.CopyFile A, ToPth$ & ZFfn_Fn(A), OvrWrt
'End Sub
'
''Sub FfnDlt(Ffn)
''If FfnIsExist(Ffn) Then Kill Ffn
''End Sub
''Function FfnExt$(Ffn)
''Dim P%: P = InStrRev(Ffn, ".")
''If P = 0 Then Exit Function
''FfnExt = Mid(Ffn, P)
''End Function
''Function FfnFdr$(Ffn)
''FfnFdr = PthFdr(FfnPth(Ffn))
''End Function
'Private Property Get ZFfn_Fn$(A)
'Dim P%: P = InStrRev(A, "\")
'If P = 0 Then ZFfn_Fn = A: Exit Property
'ZFfn_Fn = Mid(A, P + 1)
'End Property
'
'Private Property Get ZFfn_Fnn$(A)
'ZFfn_Fnn = ZFfn_RmvExt(ZFfn_Fn(A))
'End Property
'
''Function FfnIsExist(Ffn) As Boolean
''FfnIsExist = Fso.FileExists(Ffn)
''End Function
'Private Property Get ZFfn_Pth$(A)
'Dim P%: P = InStrRev(A, "\")
'If P = 0 Then Exit Property
'ZFfn_Pth = Left(A, P)
'End Property
'
'Private Property Get ZFfn_RmvExt$(A)
'Dim P%: P = InStrRev(A, ".")
'If P = 0 Then ZFfn_RmvExt = Left(A, P): Exit Property
'ZFfn_RmvExt = Left(A, P - 1)
'End Property
'
'Private Property Get ZFmtQQ$(QQVbl$, ParamArray Ap())
'Dim O$: O = Replace(QQVbl, "|", vbCrLf)
'Dim Av(): Av = Ap
'Dim I
'For Each I In Av
'    O = Replace(O, "?", I, Count:=1)
'Next
'ZFmtQQ = O
'End Property
'
'Private Property Get ZFso() As FileSystemObject
'Set ZFso = New FileSystemObject
'End Property
'
'Private Property Get ZFstChr$(A)
'ZFstChr = Left(A, 1)
'End Property
'
'Private Sub ZFt_RmvFst4Lines(Ft$)
'Dim A$: A = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
'Dim B$: B = Left(A, 55)
'Dim C$: C = Mid(A, 56)
'Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
'If B <> B1 Then Stop
'Fso.CreateTextFile(Ft, True).Write C
'End Sub
'
'Private Property Get ZHasPfx(S, Pfx$) As Boolean
'ZHasPfx = Left(S, Len(Pfx)) = Pfx
'End Property
'
'Private Property Get ZHasSubStr(A, SubStr$) As Boolean
'ZHasSubStr = InStr(A, SubStr) > 0
'End Property
'
'Private Property Get ZHdr$(W1%, W2%)
'Dim H1$: H1 = ZStr_Dup("-", W1 + 2)
'Dim H2$: H2 = ZStr_Dup("-", W2 + 2)
'ZHdr = "|" + H1 + "|" + H2 + "|"
'End Property
'
'Private Property Get ZIsNothing(A) As Boolean
'ZIsNothing = TypeName(A) = "Nothing"
'End Property
'
'Private Property Get ZIs_AllRemarked(Md As CodeModule) As Boolean
'Dim J%, L$
'For J = 1 To Md.CountOfLines
'    If Left(Md.Lines(J, 1), 1) <> "'" Then Exit Property
'Next
'ZIs_AllRemarked = True
'End Property
'
'Private Property Get ZIs_Dic_AllKeyIsNm(A As Dictionary) As Boolean
'Dim K
'For Each K In A.Keys
'    If Not ZIs_Nm(K) Then Exit Property
'Next
'ZIs_Dic_AllKeyIsNm = True
'End Property
'
'Private Property Get ZIs_Dic_AllValIsStr(A As Dictionary) As Boolean
'Dim K
'For Each K In A.Keys
'    If Not ZIs_Str(A(K)) Then Exit Property
'Next
'ZIs_Dic_AllValIsStr = True
'End Property
'
'Private Property Get ZIs_Digit(A) As Boolean
'ZIs_Digit = "0" <= A And A <= "9"
'End Property
'
'Private Property Get ZIs_Emp(V) As Boolean
'ZIs_Emp = True
'If IsMissing(V) Then Exit Property
'If ZIs_Nothing(V) Then Exit Property
'If IsEmpty(V) Then Exit Property
'If ZIs_Str(V) Then
'   If V = "" Then Exit Property
'End If
'If IsArray(V) Then
'   If ZSz(V) = 0 Then Exit Property
'End If
'ZIs_Emp = False
'End Property
'Private Property Get ZLin_ShiftMthTy$(O$)
'Dim A$
'A = "Property Get": If ZIs_Pfx(O, A) Then ZLin_ShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Property Let": If ZIs_Pfx(O, A) Then ZLin_ShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Property Set": If ZIs_Pfx(O, A) Then ZLin_ShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Function":     If ZIs_Pfx(O, A) Then ZLin_ShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Sub":          If ZIs_Pfx(O, A) Then ZLin_ShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'End Property
'Private Property Get ZIs_FunTy(A$) As Boolean
'Select Case A
'Case "Property", "Sub", "Function": ZIs_FunTy = True
'End Select
'End Property
'
'Private Property Get ZIs_Letter(A) As Boolean
'Dim C1$: C1 = UCase(A)
'ZIs_Letter = ("A" <= C1 And C1 <= "Z")
'End Property
'
'Private Property Get ZIs_Md_Exist_InPj(MdNm$, Pj As VBProject) As Boolean
'Dim I, Cmp As VBComponent
'For Each I In Pj.VBComponents
'    Set Cmp = I
'    If Cmp.Name = MdNm Then ZIs_Md_Exist_InPj = True: Exit Property
'Next
'End Property
'
'Private Property Get ZIs_Nm(A) As Boolean
'If Not ZIs_Letter(ZFstChr(A)) Then Exit Property
'Dim L%: L = Len(A)
'If L > 64 Then Exit Property
'Dim J%
'For J = 2 To L
'   If Not ZIs_NmChr(Mid(A, J, 1)) Then Exit Property
'Next
'ZIs_Nm = True
'End Property
'
'Private Property Get ZIs_NmChr(A$) As Boolean
'ZIs_NmChr = True
'If ZIs_Letter(A) Then Exit Property
'If A = "_" Then Exit Property
'If ZIs_Digit(A) Then Exit Property
'ZIs_NmChr = False
'End Property
'
'Private Property Get ZIs_Nothing(A) As Boolean
'ZIs_Nothing = TypeName(A) = "Nothing"
'End Property
'
'Private Property Get ZIs_Pfx(A, Pfx$) As Boolean
'ZIs_Pfx = Left(A, Len(Pfx)) = Pfx
'End Property
'
'Private Property Get ZIs_Prim(A) As Boolean
'Select Case VarType(A)
'Case _
'   VbVarType.vbBoolean, _
'   VbVarType.vbByte, _
'   VbVarType.vbCurrency, _
'   VbVarType.vbDate, _
'   VbVarType.vbDecimal, _
'   VbVarType.vbDouble, _
'   VbVarType.vbInteger, _
'   VbVarType.vbLong, _
'   VbVarType.vbSingle, _
'   VbVarType.vbString
'   ZIs_Prim = True
'End Select
'End Property
'
'Private Property Get ZIs_Pun(C) As Boolean
'If ZIs_Letter(C) Then Exit Property
'If ZIs_Digit(C) Then Exit Property
'If C = "_" Then Exit Property
'ZIs_Pun = True
'End Property
'
'Private Property Get ZIs_Str(A) As Boolean
'ZIs_Str = VarType(A) = vbString
'End Property
'
'Private Property Get ZIs_TstSub_Lin(L$) As Boolean
'ZIs_TstSub_Lin = True
'If ZIs_Pfx(L, "Sub Tst()") Then Exit Property
'If ZIs_Pfx(L, "Private Sub Tst()") Then Exit Property
'If ZIs_Pfx(L, "Friend Sub Tst()") Then Exit Property
'If ZIs_Pfx(L, "Sub ZZ__Tst()") Then Exit Property
'If ZIs_Pfx(L, "Private Sub ZZ__Tst()") Then Exit Property
'If ZIs_Pfx(L, "Friend Sub ZZ__Tst()") Then Exit Property
'ZIs_TstSub_Lin = False
'End Property
'
'Private Property Get ZItr_Ay(A, OIntoAy)
'Dim O: O = OIntoAy: Erase O
'Dim I
'For Each I In A
'    ZPush O, I
'Next
'ZItr_Ay = O
'End Property
'
'Private Property Get ZItr_Ny(Itr) As String()
'Dim I, O$()
'For Each I In Itr
'    ZPush O, CallByName(I, "Name", VbGet)
'Next
'ZItr_Ny = O
'End Property
'
'Private Property Get ZJnCrLf$(Ay)
'ZJnCrLf = Join(Ay, vbCrLf)
'End Property
'
'Private Property Get ZLasChr$(A)
'ZLasChr = Right(A, 1)
'End Property
'
'Private Property Get ZLinMth_LCCOpt(L$, MthNm$, Lno%) As LCCOpt
'Dim A$
'Dim M$
'Dim N$
'A = ZLin_RmvMdy(L)
'M = ZLin_ShiftMthTy(A)
'If M = "" Then Exit Property
'N = ZLin_Nm(A)
'If N <> MthNm Then Exit Property
'Dim C1%, C2%
'C1 = InStr(L, MthNm)
'C2 = C1 + Len(MthNm)
'With ZLinMth_LCCOpt
'    .Som = True
'    With .LCC
'        .Lno = Lno
'        .C1 = C1
'        .C2 = C2
'    End With
'End With
'End Property
'
'Private Property Get ZLin_FunTy$(MthLin$)
'Dim A$: A = ZLin_RmvMdy(MthLin)
'Dim B$: B = ZLin_T1(A)
'Select Case B
'Case "Function", "Sub", "Property": ZLin_FunTy = B: Exit Property
'End Select
'End Property
'
'Private Property Get ZLin_Mdy$(L$)
'Dim A$
'A = "Private": If ZHasPfx(L, A) Then ZLin_Mdy = A: Exit Property
'A = "Friend":  If ZHasPfx(L, A) Then ZLin_Mdy = A: Exit Property
'A = "Public":  If ZHasPfx(L, A) Then ZLin_Mdy = A: Exit Property
'End Property
'
'Private Property Get ZLin_Nm$(A)
'Dim J%
'If Not ZIs_Letter(Left(A, 1)) Then Exit Property
'For J = 2 To Len(A)
'    If Not ZIs_NmChr(Mid(A, J, 1)) Then
'        ZLin_Nm = Left(A, J - 1)
'        Exit Property
'    End If
'Next
'ZLin_Nm = A
'End Property
'
'Private Property Get ZLin_RmvMdy$(L$)
'Dim A$
'A = "Private ": If ZHasPfx(L, A) Then ZLin_RmvMdy = ZRmvPfx(L, A): Exit Property
'A = "Friend ":  If ZHasPfx(L, A) Then ZLin_RmvMdy = ZRmvPfx(L, A): Exit Property
'A = "Public ":  If ZHasPfx(L, A) Then ZLin_RmvMdy = ZRmvPfx(L, A): Exit Property
'ZLin_RmvMdy = L
'End Property
'
'Private Property Get ZLin_T1$(L)
'Dim A$: A = LTrim(L)
'Dim P%: P = InStr(A, " ")
'If P = 0 Then ZLin_T1 = A: Exit Property
'ZLin_T1 = Left(A, P - 1)
'End Property
'
'Private Property Get ZLinesAy_Wdt%(A$())
'Dim O%, J&, M%
'For J = 0 To ZUB(A)
'   M = ZLines_Wdt(A(J))
'   If M > O Then O = M
'Next
'ZLinesAy_Wdt = O
'End Property
'
'Private Property Get ZLines_SqV(Lines$) As Variant
'ZLines_SqV = ZAySqV(ZSplitLines(Lines))
'End Property
'
'Private Property Get ZLines_TrimEnd$(A$)
'ZLines_TrimEnd = Join(ZLy_TrimEnd(ZSplitLines(A)), vbCrLf)
'End Property
'
'Private Property Get ZLines_UnderLin$(Lines)
'ZLines_UnderLin = ZStr_Dup("-", ZLines_Wdt(Lines))
'End Property
'
'Private Property Get ZLines_Wdt%(A)
'ZLines_Wdt = ZAyWdt(ZSplitLines(A))
'End Property
'
'Private Property Get ZLy_TrimEnd(Ly) As String()
'If ZSz(Ly) = 0 Then Exit Property
'Dim L$
'Dim J&
'For J = ZUB(Ly) To 0 Step -1
'    L = Trim(Ly(J))
'    If Trim(Ly(J)) <> "" Then
'        Dim O$()
'        O = Ly
'        ReDim Preserve O(J)
'        ZLy_TrimEnd = O
'        Exit Property
'    End If
'Next
'End Property
'
'Private Property Get ZMax(A, B)
'If A > B Then
'    ZMax = A
'Else
'    ZMax = B
'End If
'End Property
'
'Private Property Get ZMbrAy() As CodeModule()
'Dim O() As CodeModule, I, Cmp As VBComponent
'For Each I In ZCurPj.VBComponents
'    Set Cmp = I
'    If Cmp.Name <> "A__" And Cmp.Name <> "M_A" Then
'        ZPushObj O, Cmp.CodeModule
'    End If
'Next
'ZMbrAy = O
'End Property
'
'Private Property Get ZCurVbe_MdPjNy(MdNm$) As String()
'ZCurVbe_MdPjNy = ZVbe_MdPjNy(ZCurVbe, MdNm)
'End Property
'
'Private Property Get ZVbe_MdPjNy(A As VBE, MdNm$) As String()
'Dim I, O$()
'For Each I In ZVbe_PjAy(A)
'    If ZPj_HasCmp(ZCvPj(I), MdNm) Then
'        ZPush O, ZCvPj(I).Name
'    End If
'Next
'ZVbe_MdPjNy = O
'End Property
'Private Property Get ZMd(PjMdDotOrColonNm) As CodeModule
'Dim A$: A = PjMdDotOrColonNm
'Dim P As VBProject
'Dim MdNm$
'    Dim L%
'    L = InStr(A, ".")
'    If L = 0 Then
'        L = InStr(A, ":")
'    End If
'    If L = 0 Then
'        Set P = ZCurPj
'        MdNm = A
'    Else
'        Dim PjNm$
'        PjNm = Left(A, L - 1)
'        Set P = ZPj(PjNm)
'        MdNm = Mid(A, L + 1)
'    End If
'Set ZMd = P.VBComponents(MdNm).CodeModule
'End Property
'
'Private Sub ZMdMth_Go(Md As CodeModule, MthNm$)
'ZMd_GoLCCOpt Md, ZMdMth_LCCOpt(Md, MthNm)
'End Sub
'
'Private Property Get ZMdMth_LCCOpt(A As CodeModule, MthNm$) As LCCOpt
'Dim L%, M As LCCOpt
'For L = A.CountOfDeclarationLines + 1 To A.CountOfLines
'    M = ZLinMth_LCCOpt(A.Lines(L, 1), MthNm, L)
'    If M.Som Then
'        ZMdMth_LCCOpt.Som = True
'        ZMdMth_LCCOpt = M
'        Exit Property
'    End If
'Next
'Stop
'End Property
'
'Private Sub ZMd_Clr(A As CodeModule, Optional IsSilent As Boolean)
'With A
'    If .CountOfLines = 0 Then Exit Sub
'    If Not IsSilent Then Debug.Print ZFmtQQ("MdClr: Md(?) of lines(?) is cleared", ZMd_Nm(A), .CountOfLines)
'    .DeleteLines 1, .CountOfLines
'End With
'End Sub
'
'Private Sub ZMd_Export(A As CodeModule)
'Dim F$: F = ZMd_SrcFfn(A)
'A.Parent.Export F
'Debug.Print ZMd_Nm(A)
'End Sub
'
'Private Sub ZMd_Gen_TstSub(A As CodeModule)
'ZMd_Rmv_TstSub A
'Dim Lines$: Lines = ZMd_TstSub_BdyLines(A)
'ZMd_Rmv_EmptyLines_AtEnd A
'If Lines <> "" Then
'    A.InsertLines A.CountOfLines + 1, Lines
'End If
'End Sub
'
'Private Sub ZMd_Go(A As CodeModule)
'Cls_Win
'With A.CodePane
'    .Show
'    .Window.WindowState = vbext_ws_Maximize
'End With
'SendKeys "%WV"
'End Sub
'
'Private Sub ZMd_GoLCCOpt(Md As CodeModule, LCCOpt As LCCOpt)
'ZMd_Go Md
'With LCCOpt
'    If .Som Then
'        With .LCC
'            Md.CodePane.TopLine = .Lno
'            Md.CodePane.SetSelection .Lno, .C1, .Lno, .C2
'        End With
'    End If
'End With
'SendKeys "^{F4}"
'End Sub
'
'Private Property Get ZMd_Has_TstSub(A As CodeModule) As Boolean
'Dim I
'For Each I In ZMd_Ly(A)
'    If I = "Friend Sub ZZ__Tst()" Then ZMd_Has_TstSub = True: Exit Property
'    If I = "Sub ZZ__Tst()" Then ZMd_Has_TstSub = True: Exit Property
'Next
'End Property
'
'Private Property Get ZMd_Lines$(A As CodeModule)
'With A
'    If .CountOfLines = 0 Then Exit Function
'    ZMd_Lines = .Lines(1, .CountOfLines)
'End With
'End Property
'
'Private Property Get ZMd_Ly(A As CodeModule) As String()
'ZMd_Ly = Split(ZMd_Lines(A), vbCrLf)
'End Property
'
'Private Property Get ZMd_MthKy(A As CodeModule, Optional IsSngLinFmt As Boolean) As String()
'Dim PjNm$: PjNm = ZMd_PjNm(A)
'Dim MdNm$: MdNm = ZMd_Nm(A)
'ZMd_MthKy = ZSrc_MthKy(ZMd_Src(A), PjNm, MdNm, IsSngLinFmt)
'End Property
'
'Private Property Get ZPj_MthKy(A As VBProject, Optional IsSngLinFmt As Boolean) As String()
'Dim O$(), I
'For Each I In ZPj_MbrAy(A)
'    ZPushAy O, ZMd_MthKy(ZCvMd(I), IsSngLinFmt)
'Next
'ZPj_MthKy = O
'End Property
'
'Private Property Get ZVbe_MthKy(A As VBE, Optional IsSngLinFmt As Boolean) As String()
'Dim O$(), I
'For Each I In ZVbe_PjAy(A)
'    ZPushAy O, ZPj_MthKy(ZCvPj(I), IsSngLinFmt)
'Next
'ZVbe_MthKy = O
'End Property
'
'Private Property Get ZMd_MthNy(A As CodeModule, Optional MthNmPatn$ = ".", Optional IsNoMdNmPfx As Boolean) As String()
'Dim Ay$(): Ay = ZSrc_MthNy(ZMd_Src(A), MthNmPatn)
'If IsNoMdNmPfx Then
'    ZMd_MthNy = Ay
'Else
'    ZMd_MthNy = ZAyAddPfx(Ay, ZMd_Nm(A) & ".")
'End If
'End Property
'
'Private Property Get ZMd_MthS1S2Ay(A As CodeModule) As S1S2()
'Dim P$: P = ZMd_PjNm(A)
'Dim M$: M = ZMd_Nm(A)
'ZMd_MthS1S2Ay = ZSrc_MthS1S2Ay(ZMd_Src(A), P, M)
'End Property
'
'Private Property Get ZMd_Nm$(A As CodeModule)
'ZMd_Nm = A.Parent.Name
'End Property
'
'Private Property Get ZMd_Pj(A As CodeModule) As VBProject
'Set ZMd_Pj = A.Parent.Collection.Parent
'End Property
'
'Private Property Get ZMd_PjNm$(A As CodeModule)
'ZMd_PjNm = ZMd_Pj(A).Name
'End Property
'
'Private Property Get ZMd_Rmk(Md As CodeModule) As Boolean
'Debug.Print "Rmk " & Md.Parent.Name,
'If ZIs_AllRemarked(Md) Then
'    Debug.Print " No need"
'    Exit Property
'End If
'Debug.Print "<============= is remarked"
'Dim J%
'For J = 1 To Md.CountOfLines
'    Md.ReplaceLine J, "'" & Md.Lines(J, 1)
'Next
'ZMd_Rmk = True
'End Property
'
'Private Sub ZMd_Rmv_EmptyLines_AtEnd(A As CodeModule)
'Dim J%
'While A.CountOfLines > 1
'    J = J + 1
'    If J > 10000 Then Stop
'    If Trim(A.Lines(A.CountOfLines, 1)) <> "" Then Exit Sub
'    A.DeleteLines A.CountOfLines, 1
'Wend
'End Sub
'
'Private Sub ZMd_Rmv_TstSub(A As CodeModule)
'Dim L&, N&
'L = ZMd_TstSub_Lno(A)
'If L = 0 Then Exit Sub
'Dim Fnd As Boolean, J%
'For J = L + 1 To A.CountOfLines
'    If ZIs_Pfx(A.Lines(J, 1), "End Sub") Then
'        N = J - L + 1
'        Fnd = True
'        Exit For
'    End If
'Next
'If Not Fnd Then Stop
'A.DeleteLines L, N
'End Sub
'
'Private Property Get ZMd_Src(A As CodeModule) As String()
'ZMd_Src = ZMd_Ly(A)
'End Property
'
'Private Property Get ZMd_SrcExt$(A As CodeModule)
'Dim O$
'Select Case A.Parent.Type
'Case vbext_ct_ClassModule: O = ".cls"
'Case vbext_ct_Document: O = ".cls"
'Case vbext_ct_StdModule: O = ".bas"
'Case vbext_ct_MSForm: O = ".cls"
'Case Else: Err.Raise 1, , "ZMd_SrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
'End Select
'ZMd_SrcExt = O
'End Property
'
'Private Property Get ZMd_SrcFfn$(A As CodeModule)
'ZMd_SrcFfn = ZPj_SrcPth(ZMd_Pj(A)) & ZMd_SrcFn(A)
'End Property
'
'Private Property Get ZMd_SrcFn$(A As CodeModule)
'ZMd_SrcFn = ZMd_Nm(A) & ZMd_SrcExt(A)
'End Property
'
'Private Sub ZMd_Srt(A As CodeModule)
'If ZMd_Nm(A) = "M_Tool" And ZMd_PjNm(A) = "QTool" Then
'    Exit Sub
'End If
'Dim Nm$: Nm = ZMd_Nm(A)
'Debug.Print "Sorting: "; ZAlignL(Nm, 30); " ";
'Dim Ay(): Ay = Array("M_A")
''Skip some md
'    If ZAyHas(Ay, Nm) Then
'        Debug.Print "<<<< Skipped"
'        Exit Sub
'    End If
'Dim NewLines$: NewLines = ZMd_SrtedLines(A)
'Dim Old$: Old = ZMd_Lines(A)
''Exit if same
'    If Old = NewLines Then
'        Debug.Print "<== Same"
'        Exit Sub
'    End If
'Debug.Print "<-- Sorted";
''Delete
'    Debug.Print ZFmtQQ("<--- Deleted (?) lines", A.CountOfLines);
'    ZMd_Clr A, IsSilent:=True
''Add sorted lines
'    A.AddFromString NewLines
'    ZMd_Rmv_EmptyLines_AtEnd A
'    Debug.Print "<----Sorted Lines added...."
'End Sub
'
'Private Property Get ZMd_SrtRpt(A As CodeModule) As DCRslt
'Dim PjNm$, MdNm$
'MdNm = ZMd_Nm(A)
'PjNm = ZMd_PjNm(A)
'ZMd_SrtRpt = ZSrc_SrtRpt(ZMd_Src(A), PjNm, MdNm)
'End Property
'
'Private Property Get ZMd_SrtRptLy(A As CodeModule) As String()
'Dim PjNm$: PjNm = ZMd_PjNm(A)
'Dim MdNm$: MdNm = ZMd_Nm(A)
'ZMd_SrtRptLy = ZSrc_SrtRptLy(ZMd_Src(A), PjNm, MdNm)
'End Property
'
'Private Property Get ZMd_SrtedLines$(A As CodeModule)
'ZMd_SrtedLines = ZSrc_SrtedLines(ZMd_Src(A))
'End Property
'
'Private Property Get ZMd_TstSub_BdyLines$(A As CodeModule)
'Dim Ny$(): Ny = ZMd_ZZFun_Ny(A)
'If ZSz(Ny) = 0 Then Exit Property
'Ny = ZAySrt(Ny)
'Dim O$()
'Dim Pfx$
'If A.Parent.Type = vbext_ct_ClassModule Then
'    Pfx = "Friend "
'End If
'ZPush O, ""
'ZPush O, Pfx & "Sub ZZ__Tst()"
'ZPushAy O, Ny
'ZPush O, "End Sub"
'ZMd_TstSub_BdyLines = Join(O, vbCrLf)
'End Property
'
'Private Property Get ZMd_TstSub_Lno%(A As CodeModule)
'Dim J%
'For J = 1 To A.CountOfLines
'    If ZIs_TstSub_Lin(A.Lines(J, 1)) Then ZMd_TstSub_Lno = J: Exit Property
'Next
'End Property
'
'Private Property Get ZMd_UnRmk(Md As CodeModule) As Boolean
'Debug.Print "UnRmk " & Md.Parent.Name,
'If Not ZIs_AllRemarked(Md) Then
'    Debug.Print "No need"
'    Exit Property
'End If
'Debug.Print "<===== is unmarked"
'Dim J%, L$
'For J = 1 To Md.CountOfLines
'    L = Md.Lines(J, 1)
'    If Left(L, 1) <> "'" Then Stop
'    Md.ReplaceLine J, Mid(L, 2)
'Next
'ZMd_UnRmk = True
'End Property
'
'Private Property Get ZMd_ZZFun_Ny(A As CodeModule) As String()
'Dim J%, O$(), L$, L1$, Is_ZZFun As Boolean
'For J = 1 To A.CountOfLines
'    Is_ZZFun = True
'    L = A.Lines(J, 1)
'    Select Case True
'    Case ZIs_Pfx(L, "Private Sub ZZ_")
'        Is_ZZFun = True
'        L1 = ZRmvPfx(L, "Private Sub ")
'    Case ZIs_Pfx(L, "Sub ZZ_")
'        Is_ZZFun = True
'        L1 = ZRmvPfx(L, "Sub ")
'    Case Else:
'        Is_ZZFun = False
'    End Select
'
'    If Is_ZZFun Then
'        ZPush O, ZLin_Nm(L1)
'    End If
'Next
'ZMd_ZZFun_Ny = O
'End Property
'
'Private Sub ZMthLin_BrkAsg(A$, Optional OIsMthLin As Boolean, Optional OMdy$, Optional OMajTy$, Optional OMthNm$)
'OMdy = ZLin_Mdy(A)
'OMthNm = ""
'OMajTy = ""
'
'Dim L$
'    If OMdy = "" Then L = A Else L = ZRmvPfx(A, OMdy & " ")
'
''OMajTy
'    Dim B$
'    B = "Sub ":          If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): OMajTy = "Sub"
'    B = "Function ":     If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): OMajTy = "Fun"
'    B = "Property Get ": If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): OMajTy = "Prp"
'    B = "Property Let ": If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): OMajTy = "Prp"
'    B = "Property Set ": If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): OMajTy = "Prp"
'    If OMajTy = "" Then
'        OIsMthLin = False
'        Exit Sub
'    End If
'OMthNm = ZLin_Nm(L)
'OIsMthLin = True
'End Sub
'
'Private Property Get ZMthLin_MthKey$(A$, Optional PjNm$ = "Pj", Optional MdNm$ = "Md", Optional IsSngLinFmt As Boolean)
'Dim M$ 'Mdy
'Dim T$ 'MthTy {Sub Fun Prp}
'Dim N$ 'Name
'Dim P% 'Priority
'    M = ZLin_Mdy(A)
'    Dim L$
'    If M = "" Then L = A Else L = ZRmvPfx(A, M & " ")
'    Dim B$
'    B = "Sub ":          If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): T = "Sub"
'    B = "Function ":     If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): T = "Fun"
'    B = "Property Get ": If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): T = "Prp"
'    B = "Property Let ": If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): T = "Prp"
'    B = "Property Set ": If ZHasPfx(L, B) Then L = ZRmvPfx(L, B): T = "Prp"
'    If T = "" Then Stop
'    N = ZLin_Nm(L)
'If ZIs_Pfx(N, "Init") And T = "Get" And M = "Friend" Then
'    P = 1
'ElseIf T = "Prp" And (M = "" Or M = "Public") Then
'    P = 2
'ElseIf ZHasSubStr(N, "__") Then
'    P = 4
'ElseIf N = "ZZ__Tst" Then
'    P = 9
'ElseIf ZIs_Pfx(N, "ZZ_") Then
'    P = 8
'ElseIf M = "Private" Then
'    P = 5
'Else
'    P = 3
'End If
'Dim F$
'F = IIf(IsSngLinFmt, "?:?:?:?:?:?", "?:?|?:?|?:?")
'ZMthLin_MthKey = ZFmtQQ(F, PjNm, MdNm, P, N, T, M)
'End Property
'
'Private Property Get ZMthLin_MthNm$(A$)
'Dim N$ 'Name
'    ZMthLin_BrkAsg A, _
'        OMthNm:=N
'ZMthLin_MthNm = N
'End Property
'
'Private Function ZNewWb() As Workbook
'ZXls.Workbook.Add
'End Function
'
'Private Sub ZOyDo(Oy, DoMthNm$)
'Dim O
'For Each O In Oy
'    Run DoMthNm, O ' DoMthNm call be like a Excel.Address (eg, A1, XX1)
'Next
'End Sub
'
'Private Property Get ZOy_Ny(Oy) As String()
'Dim O$(): If ZSz(Oy) = 0 Then Exit Property
'Dim I
'For Each I In Oy
'    ZPush O, CallByName(I, "Name", VbGet)
'Next
'ZOy_Ny = O
'End Property
'
'Private Property Get ZPj(PjNm$) As VBProject
'Set ZPj = ZCurVbe.VBProjects(PjNm)
'End Property
'
'Private Sub ZPjMdMthDotNm_BrkAsg(A$, OMd As CodeModule, OMthNm$)
'Dim P$, M$
'    ZDotDotNm_BrkAsg A, _
'        P, M, OMthNm
'Dim Pj As VBProject
'    If P = "" Then
'        Set Pj = ZCurPj
'    Else
'        Set Pj = ZPj(P)
'    End If
'Set OMd = ZPj_Md(Pj, M)
'End Sub
'
'Private Sub ZPj_AddRf(A As VBProject, RfNm$)
'Dim RfFfn$: RfFfn = ZRfNm_RfFfn(RfNm)
'If RfFfn = "" Then Stop
'Dim F$: F = ZPj_Ffn(A)
'If F = "" Then Exit Sub
'If F = RfFfn Then Exit Sub
'If ZPj_HasRfNm(A, RfNm) Then Exit Sub
'A.References.AddFromFile RfFfn
'ZPj_Sav A
'End Sub
'
'Private Sub ZPj_Add_Cls(A As VBProject, Nm$)
'ZPj_Add_Mbr A, Nm, vbext_ct_ClassModule
'End Sub
'
'Private Sub ZPj_Add_Mbr(A As VBProject, Nm$, Ty As vbext_ComponentType, Optional IsGoMbr As Boolean)
'If ZPj_HasCmp(A, Nm) Then
'    MsgBox ZFmtQQ("Cmp(?) exist in CurPj(?)", Nm, ZCurPjNm), , "M_A.ZAddMbr"
'    Exit Sub
'End If
'Dim Cmp As VBComponent
'Set Cmp = A.VBComponents.Add(Ty)
'Cmp.Name = Nm
'Cmp.CodeModule.AddFromString "Option Explicit"
'If IsGoMbr Then Go_Mbr Nm
'End Sub
'
'Private Property Get ZPj_ClsNy_With_TstSub(A As VBProject) As String()
'Dim I As VBComponent
'Dim O$()
'For Each I In A.VBComponents
'    If I.Type = vbext_ct_ClassModule Then
'        If ZMd_Has_TstSub(I.CodeModule) Then
'            ZPush O, I.Name
'        End If
'    End If
'Next
'ZPj_ClsNy_With_TstSub = O
'End Property
'
'Private Property Get ZPj_Cmp(A As VBProject, Nm) As VBComponent
'Set ZPj_Cmp = A.VBComponents(CStr(Nm))
'End Property
'
'Private Sub ZPj_Compile(A As VBProject)
'ZPj_Go A
'SendKeys "%D{Enter}"
'End Sub
'
'Private Sub ZPj_Ens_Cls(A As VBProject, ClsNm$, ClsCxt$)
'ZPj_Ens_Cmp A, ClsNm, vbext_ct_StdModule, ClsCxt
'End Sub
'
'Private Sub ZPj_Ens_Cmp(A As VBProject, Nm$, Ty As vbext_ComponentType, Cxt$)
'If Not ZPj_HasCmp(A, Nm) Then
'    Dim Cmp As VBComponent
'    Set Cmp = A.VBComponents.Add(Ty)
'    Cmp.Name = Nm
'    Cmp.CodeModule.InsertLines 1, Cxt
'    Debug.Print ZFmtQQ("ZPj_Ens_Cmp: Md(?) of Ty(?) with Cxt-Len(?) is added in Pj(?) <===================================", Nm, ZCmpTy_Nm(Ty), Len(Cxt), A.Name)
'    Exit Sub
'End If
'Dim Md As CodeModule
'    Set Md = ZPj_Md(A, Nm)
'If ZMd_Lines(Md) = Cxt Then
'    Debug.Print ZFmtQQ("ZPj_Ens_Cmp: Md(?) of Ty(?) with Cxt-Len(?) is same as in Pj(?)", Nm, ZCmpTy_Nm(Ty), Len(Cxt), A.Name)
'    Exit Sub
'End If
'ZMd_Clr Md
'Md.InsertLines 1, Cxt
'Debug.Print ZFmtQQ("ZPj_Ens_Cmp: Md(?) of Ty(?) with Cxt-Len(?) is replaced as in Pj(?)<-----------------", Nm, ZCmpTy_Nm(Ty), Len(Cxt), A.Name)
'End Sub
'
'Private Sub ZPj_Ens_Md(A As VBProject, MdNm$, MdCxt$)
'ZPj_Ens_Cmp A, MdNm, vbext_ct_StdModule, MdCxt
'End Sub
'
'Private Sub ZPj_Export(A As VBProject)
'Dim P$: P = ZPj_SrcPth(A)
'If P = "" Then Exit Sub
'ZPth_ClrFil P 'Clr SrcPth ---
'ZFfn_CpyToPth A.Filename, P, OvrWrt:=True
'Dim I, Ay() As CodeModule
'Ay = ZPj_MbrAy(A)
'If ZSz(Ay) = 0 Then Exit Sub
'For Each I In Ay
'    ZMd_Export ZCvMd(I)  'Exp each md --
'Next
'ZAyWrt ZPj_RfLy(A), ZPj_RfCfgFfn(A) 'Exp rf -----
'End Sub
'
'Private Property Get ZPj_Ffn$(A As VBProject)
'On Error Resume Next
'ZPj_Ffn = A.Filename
'End Property
'
'Private Property Get ZPj_FstMd(A As VBProject) As CodeModule
'Dim Cmp As VBComponent, O$()
'For Each Cmp In A.VBComponents
'    If Cmp.Type = vbext_ct_ClassModule Or Cmp.Type = vbext_ct_StdModule Then
'        Set ZPj_FstMd = Cmp.CodeModule
'        Exit Property
'    End If
'Next
'End Property
'
'Private Sub ZPj_Gen_TstClass(A As VBProject)
'If ZPj_HasCmp(A, "Tst") Then
'    ZCmp_Rmv ZPj_Cmp(A, "Tst")
'End If
'ZPj_Add_Cls A, "Tst"
'ZPj_Md(A, "Tst").AddFromString ZPj_TstClass_Bdy(A)
'End Sub
'
'Private Sub ZPj_Gen_TstSub(A As VBProject)
'Dim Ny$(): Ny = ZPj_Md_and_Cls_Ny(A)
'Dim N, M As CodeModule
'For Each N In Ny
'    Set M = A.VBComponents(N).CodeModule
'    ZMd_Gen_TstSub M
'Next
'End Sub
'
'Private Sub ZPj_Go(A As VBProject)
'Cls_Win
'Dim Md As CodeModule
'Set Md = ZPj_FstMd(A)
'If ZIsNothing(Md) Then Exit Sub
'Debug.Print ZMd_Nm(Md)
'Md.CodePane.Show
'SendKeys "%WV" ' Window SplitVertical
'End Sub
'
'Private Property Get ZPj_HasCmp(A As VBProject, Nm$) As Boolean
'Dim Cmp As VBComponent
'For Each Cmp In A.VBComponents
'    If Cmp.Name = Nm Then ZPj_HasCmp = True: Exit Property
'Next
'End Property
'
'Private Property Get ZPj_HasMdNm(A As VBProject, MdNm$) As Boolean
'Dim Cmp As VBComponent, I
'For Each I In A.VBComponents
'    If I.Cmp = MdNm Then ZPj_HasMdNm = True: Exit Property
'Next
'End Property
'
'Private Property Get ZPj_HasRfNm(A As VBProject, RfNm$) As Boolean
'Dim I, R As Reference
'For Each I In A.References
'    Set R = I
'    If R.Name = RfNm Then ZPj_HasRfNm = True: Exit Property
'Next
'End Property
'
'Private Property Get ZPj_MbrAy(A As VBProject, Optional MbrNmPatn$ = ".") As CodeModule()
'Dim O() As CodeModule
'Dim Cmp As VBComponent
'Dim R As RegExp: If MbrNmPatn <> "." Then Set R = ZRe(MbrNmPatn)
'For Each Cmp In A.VBComponents
'    Select Case Cmp.Type
'    Case vbext_ct_StdModule, vbext_ct_ClassModule
'        If MbrNmPatn = "." Then
'            ZPushObj O, Cmp.CodeModule
'        Else
'            If ZReTst(R, Cmp.Name) Then
'                ZPushObj O, Cmp.CodeModule
'            End If
'        End If
'    End Select
'Next
'ZPj_MbrAy = O
'End Property
'
'Private Property Get ZPj_MbrNy(A As VBProject, Optional MbrNmPatn$ = ".") As String()
'ZPj_MbrNy = ZOy_Ny(ZPj_MbrAy(A, MbrNmPatn))
'End Property
'
'Private Property Get ZPj_MbrAyLik(A As VBProject, MdLikNm$) As CodeModule()
'Dim Cmp As VBComponent, O() As CodeModule
'For Each Cmp In A.VBComponents
'    If Cmp.Type = vbext_ct_ClassModule Or Cmp.Type = vbext_ct_StdModule Then
'        If Cmp.Name Like MdLikNm Then
'            ZPushObj O, Cmp
'        End If
'    End If
'Next
'ZPj_MbrAyLik = O
'End Property
'
'Private Property Get ZPj_MbrNyLik(A As VBProject, MdLikNm$) As String()
'ZPj_MbrNyLik = ZOy_Ny(ZPj_MbrAyLik(A, MdLikNm))
'End Property
'
'Private Property Get ZPj_Md(A As VBProject, Nm) As CodeModule
'Set ZPj_Md = ZPj_Cmp(A, Nm).CodeModule
'End Property
'
'Private Property Get ZPj_MdNy_With_TstSub(A As VBProject) As String()
'Dim I As VBComponent
'Dim O$()
'For Each I In A.VBComponents
'    If I.Type = vbext_ct_StdModule Then
'        If ZMd_Has_TstSub(I.CodeModule) Then
'            ZPush O, I.Name
'        End If
'    End If
'Next
'ZPj_MdNy_With_TstSub = O
'End Property
'
'Private Property Get ZPj_MdSrtRpt(A As VBProject) As MdSrtRpt
''SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
'Dim Ay() As CodeModule: Ay = ZPj_MbrAy(A)
'Dim Ny$(): Ny = ZOy_Ny(Ay)
'Dim LyAy()
'Dim IsSam() As Boolean
'    Dim J%, R As DCRslt
'    For J = 0 To ZUB(Ay)
'        R = ZMd_SrtRpt(Ay(J))
'        ZPush LyAy, ZDCRslt_Ly(R)
'        ZPush IsSam, ZDCRslt_IsSam(R)
'    Next
'With ZPj_MdSrtRpt
'    Set .RptDic = ZAyPair_Dic(Ny, LyAy)
'    .MdNy = ZPj_MdSrtRpt_1(Ny, IsSam)
'End With
'End Property
'
'Private Property Get ZPj_MdSrtRpt_1(MdNy$(), IsSam() As Boolean) As String()
'Dim O$(), J%
'For J = 0 To ZUB(MdNy)
'    ZPush O, ZAlignL(MdNy(J), 30) & " " & IsSam(J)
'Next
'ZPj_MdSrtRpt_1 = O
'End Property
'
'Private Property Get ZPj_Md_and_Cls_Ny(A As VBProject) As String()
'Dim O$(), Cmp As VBComponent
'For Each Cmp In A.VBComponents
'    If Cmp.Type = vbext_ct_StdModule Or Cmp.Type = vbext_ct_ClassModule Then
'        ZPush O, Cmp.Name
'    End If
'Next
'ZPj_Md_and_Cls_Ny = O
'End Property
'
'Private Property Get ZPj_MthNy(A As VBProject, Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".") As String()
'Dim Ay() As CodeModule: Ay = ZPj_MbrAy(A, MbrNmPatn)
'If ZSz(Ay) = 0 Then Exit Property
'Dim I, O$()
'For Each I In Ay
'    ZPushAy O, ZMd_MthNy(ZCvMd(I), MthNmPatn)
'Next
'O = ZAyAddPfx(O, A.Name & ".")
'ZPj_MthNy = O
'End Property
'
'Private Property Get ZPj_MthS1S2Ay(A As VBProject) As S1S2()
'Dim I
'Dim Ay() As CodeModule: Ay = ZPj_MbrAy(A)
'Dim O() As S1S2
'Dim M As CodeModule
'For Each I In Ay
'    Set M = I
'    O = ZS1S2Ay_Add(O, ZMd_MthS1S2Ay(M))
'Next
'ZPj_MthS1S2Ay = O
'End Property
'
'Private Property Get ZPj_RfAy(A As VBProject) As Reference()
'ZPj_RfAy = ZItr_Ay(A.References, ZEmp_RfAy)
'End Property
'
'Private Property Get ZPj_RfCfgFfn(A As VBProject)
'ZPj_RfCfgFfn = ZPj_SrcPth(A) & "PjRf.Cfg"
'End Property
'
'Private Property Get ZPj_RfLy(A As VBProject) As String()
'Dim RfAy() As Reference
'    RfAy = ZPj_RfAy(A)
'Dim O$()
'Dim Ny$(): Ny = ZOy_Ny(RfAy)
'Ny = ZAyAlignL(Ny)
'Dim J%
'For J = 0 To ZUB(Ny)
'    ZPush O, Ny(J) & " " & ZRf_Ffn(RfAy(J))
'Next
'ZPj_RfLy = O
'End Property
'
'Private Sub ZPj_Sav(A As VBProject)
'ZPj_Go A
'SendKeys "^S"
'End Sub
'
'Private Property Get ZPj_SrcPth(A As VBProject)
'Dim Ffn$: Ffn = ZPj_Ffn(A)
'If Ffn = "" Then Exit Property
'Dim Fn$: Fn = ZFfn_Fn(Ffn)
'Dim P$: P = ZFfn_Pth(A.Filename)
'If P = "" Then Exit Property
'Dim O$:
'O = P & "Src\": ZPth_Ens O
'O = O & Fn & "\":                  ZPth_Ens O
'ZPj_SrcPth = O
'End Property
'
'Private Sub ZPj_SrcPthBrw(A As VBProject)
'ZPth_Brw ZPj_SrcPth(A)
'End Sub
'
'Private Sub ZPj_Srt(A As VBProject)
'If A.Name = "QTool" Then Exit Sub
'Dim I
'For Each I In ZPj_Md_and_Cls_Ny(A)
'    ZMd_Srt ZPj_Md(A, I)
'Next
'End Sub
'
'Private Property Get ZPj_SrtRptLy(A As VBProject) As String()
'Dim Ay() As CodeModule: Ay = ZPj_MbrAy(A)
'Dim O$(), I, M As CodeModule
'For Each I In Ay
'    Set M = I
'    ZPushAy O, ZMd_SrtRptLy(M)
'Next
'ZPj_SrtRptLy = O
'End Property
'
'Private Function ZPj_SrtRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
'Dim A1 As MdSrtRpt
'A1 = ZPj_MdSrtRpt(A)
'Dim O As Workbook: Set O = ZDic_Wb(A1.RptDic)
'Dim Ws As Worksheet
'Set Ws = ZWb_AddWs(O, "Md Idx")
''Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
''LoCol_LnkWs Lo, "Md"
''If Vis Then WbVis O
'Set ZPj_SrtRptWb = O
'Stop '
'End Function
'
'Private Property Get ZPj_TstClass_Bdy$(A As VBProject)
'Dim N1$() ' All Class Ny with 'Friend Sub ZZ__Tst' method
'Dim N2$()
'Dim A1$, A2$
'Const Q1$ = "Sub ?()|Dim A As New ?: A.ZZ__Tst|End Sub"
'Const Q2$ = "Sub ?()|#.?.ZZ__Tst|End Sub"
'N1 = ZPj_ClsNy_With_TstSub(A)
'A1 = ZSeed_Expand(Q1, N1)
'N2 = ZPj_MdNy_With_TstSub(A)
'A2 = Replace(ZSeed_Expand(Q2, N2), "#", A.Name)
'ZPj_TstClass_Bdy = A1 & vbCrLf & A2
'End Property
'
''Function FfnRplExt$(Ffn, NewExt)
''FfnRplExt = FfnRmvExt(Ffn) & NewExt
''End Function
''Sub FtBrw(Ft)
'''Shell "code.cmd """ & Ft & """", vbHide
''Shell "notepad.exe """ & Ft & """", vbMaximizedFocus
''End Sub
''Function FtLines$(Ft)
''FtLines = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
''End Function
''Function FtDic(Ft) As Dictionary
''Set FtDic = Ly(FtLy(Ft)).Dic
''End Function
''Function FtLy(Ft) As String()
''Dim F%: F = FtOpnInp(Ft)
''Dim L$, O$()
''While Not EOF(F)
''    Line Input #F, L
''    Push O, L
''Wend
''Close #F
''FtLy = O
''End Function
''Function FtOpnApp%(Ft)
''Dim O%: O = FreeFile(1)
''Open Ft For Append As #O
''FtOpnApp = O
''End Function
''Function FtOpnInp%(Ft)
''Dim O%: O = FreeFile(1)
''Open Ft For Input As #O
''FtOpnInp = O
''End Function
''Function FtOpnOup%(Ft)
''Dim O%: O = FreeFile(1)
''Open Ft For Output As #O
''FtOpnOup = O
''End Function
'Private Sub ZPth_Brw(P)
'Shell "Explorer """ & P & """", vbMaximizedFocus
'End Sub
'
'Private Sub ZPth_ClrFil(A)
'If Not ZPth_IsExist(A) Then Exit Sub
'Dim Ay$(): Ay = ZPth_FfnAy(A)
'Dim F
'On Error Resume Next
'For Each F In Ay
'   Kill F
'Next
'End Sub
'
'Private Sub ZPth_Ens(P$)
'If Fso.FolderExists(P) Then Exit Sub
'MkDir P
'End Sub
'
''Function PthEntAy(A, Optional FilSpec$ = "*.*", Optional Atr As FileAttribute, Optional IsRecursive As Boolean) As String()
''If Not IsRecursive Then
''    PthEntAy = AyAdd(PthSubPthAy(A), PthFfnAy(A, FilSpec, Atr))
''    Exit Function
''End If
''Erase O
''PthPushEntAyR A
''PthEntAy = O
''Erase O
''End Function
''Function PthFdr$(A$)
''Ass PthHasPthSfx(A)
''Dim P$: P = RmvLasChr(A)
''PthFdr = TakAftRev(A, "\")
''End Function
'Private Property Get ZPth_FfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
'ZPth_FfnAy = ZAyAddPfx(ZPth_FnAy(A, Spec, Atr), A)
'End Property
'
'Private Property Get ZPth_FnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
'ZAss ZPth_IsExist(A)
'Dim O$()
'Dim M$
'M = Dir(A & Spec)
'If Atr = 0 Then
'    While M <> ""
'       ZPush O, M
'       M = Dir
'    Wend
'    ZPth_FnAy = O
'End If
'ZAss ZPth_HasPthSfx(A)
'While M <> ""
'    If GetAttr(A & M) And Atr Then
'        ZPush O, M
'    End If
'    M = Dir
'Wend
'ZPth_FnAy = O
'End Property
'
'Private Property Get ZPth_HasPthSfx(A) As Boolean
'ZPth_HasPthSfx = ZLasChr(A) = "\"
'End Property
'
'Private Property Get ZPth_IsExist(A) As Boolean
'ZAss ZPth_HasPthSfx(A)
'ZPth_IsExist = Fso.FolderExists(A)
'End Property
'
'Private Sub ZPush(O, M)
'Dim N&
'    N = ZSz(O)
'ReDim Preserve O(N)
'If IsObject(M) Then
'    Set O(N) = M
'Else
'    O(N) = M
'End If
'End Sub
'
'Private Sub ZPushAy(OAy, Ay)
'If ZSz(Ay) = 0 Then Exit Sub
'Dim I
'For Each I In Ay
'    ZPush OAy, I
'Next
'End Sub
'
'Private Sub ZPushNoDup(O, M)
'If Not ZAyHas(O, M) Then ZPush O, M
'End Sub
'
'Private Sub ZPushNonEmp(O, M)
'If ZIs_Emp(M) Then Exit Sub
'ZPush O, M
'End Sub
'
'Private Sub ZPushObj(O, M)
'If Not IsObject(M) Then Stop
'Dim N&
'    N = ZSz(O)
'ReDim Preserve O(N)
'Set O(N) = M
'End Sub
'
''===================
'Private Property Get ZRe(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
'Dim O As New RegExp
'With O
'   .Pattern = Patn
'   .MultiLine = MultiLine
'   .IgnoreCase = IgnoreCase
'   .Global = IsGlobal
'End With
'Set ZRe = O
'End Property
'
'Private Property Get ZReMatch(A As RegExp, S) As MatchCollection
'Set ZReMatch = A.Execute(S)
'End Property
'
'Private Property Get ZReRpl$(A As RegExp, S, R$)
'ZReRpl = A.Replace(S, R)
'End Property
'
'Private Property Get ZReTst(A As RegExp, S) As Boolean
'ZReTst = A.Test(S)
'End Property
'
'Private Property Get ZRfNm_RfFfn$(RfNm$)
'Dim Ay() As VBProject: Ay = ZCurVbe_PjAy
'Dim M As VBProject, I
'For Each I In Ay
'    Set M = I
'    If M.Name = RfNm Then ZRfNm_RfFfn = M.Filename: Exit Property
'Next
'End Property
'
'Private Property Get ZRf_Ffn$(A As Reference)
'On Error Resume Next
'ZRf_Ffn = A.FullPath
'End Property
'
'Private Property Get ZRmvPfx$(A, Pfx$)
'Dim L%: L = Len(Pfx)
'If Left(A, L) = Pfx Then
'    ZRmvPfx = Mid(A, L + 1)
'Else
'    ZRmvPfx = A
'End If
'End Property
'
'Private Property Get ZRpl_DblSpc$(A)
'Dim O$: O = Trim(A)
'Dim J&
'While ZHasSubStr(O, "  ")
'    J = J + 1: If J > 10000 Then Stop
'    O = Replace(O, "  ", " ")
'Wend
'ZRpl_DblSpc = O
'End Property
'
'Private Property Get ZRpl_Pun$(A)
'Dim O$(), J&, L&, C$
'L = Len(A)
'If L = 0 Then Exit Property
'ReDim O(L - 1)
'For J = 1 To L
'    C = Mid(A, J, 1)
'    If ZIs_Pun(C) Then
'        O(J - 1) = " "
'    Else
'        O(J - 1) = C
'    End If
'Next
'ZRpl_Pun = Join(O, "")
'End Property
'
'Private Property Get ZRpl_VBar$(A)
'ZRpl_VBar = Replace(A, "|", vbCrLf)
'End Property
'
'Private Property Get ZS1S2(S1$, S2$) As S1S2
'ZS1S2.S1 = S1
'ZS1S2.S2 = S2
'End Property
'
'Private Property Get ZS1S2Ay_Add(A() As S1S2, B() As S1S2) As S1S2()
'Dim O() As S1S2
'Dim J&
'O = A
'For J = 0 To ZS1S2_UB(B)
'    ZS1S2_Push O, B(J)
'Next
'ZS1S2Ay_Add = O
'End Property
'
'Private Sub ZS1S2Ay_Brw(A() As S1S2)
'ZAyBrw ZS1S2Ay_FmtLy(A)
'End Sub
'
'Private Property Get ZS1S2Ay_Dic(A() As S1S2) As Dictionary
'Dim J&, O As New Dictionary
'For J = 0 To ZS1S2_UB(A)
'    O.Add A(J).S1, A(J).S2
'Next
'Set ZS1S2Ay_Dic = O
'End Property
'
'Private Property Get ZS1S2Ay_FmtLy(A() As S1S2) As String()
'Dim W1%: W1 = ZS1S2Ay_S1LinesWdt(A)
'Dim W2%: W2 = ZS1S2Ay_S2LinesWdt(A)
'Dim H$: H = ZHdr(W1, W2)
'ZS1S2Ay_FmtLy = ZS1S2Ay_LinesLinesLy(A, H, W1, W2)
'End Property
'
'Private Property Get ZS1S2Ay_LinesLinesLy(A() As S1S2, H$, W1%, W2%) As String()
'Dim O$(), I&
'ZPush O, H
'For I = 0 To ZS1S2_UB(A)
'   ZPushAy O, ZS1S2_Ly(A(I), W1, W2)
'   ZPush O, H
'Next
'ZS1S2Ay_LinesLinesLy = O
'End Property
'
'Private Property Get ZS1S2Ay_S1LinesWdt%(A() As S1S2)
'ZS1S2Ay_S1LinesWdt = ZLinesAy_Wdt(ZS1S2Ay_Sy1(A))
'End Property
'
'Private Property Get ZS1S2Ay_S2LinesWdt%(A() As S1S2)
'ZS1S2Ay_S2LinesWdt = ZLinesAy_Wdt(ZS1S2Ay_Sy2(A))
'End Property
'
'Private Property Get ZS1S2Ay_Sy1(A() As S1S2) As String()
'Dim O$(), J&
'For J = 0 To ZS1S2_UB(A)
'   ZPush O, A(J).S1
'Next
'ZS1S2Ay_Sy1 = O
'End Property
'
'Private Property Get ZS1S2Ay_Sy2(A() As S1S2) As String()
'Dim O$(), J&
'For J = 0 To ZS1S2_UB(A)
'   ZPush O, A(J).S2
'Next
'ZS1S2Ay_Sy2 = O
'End Property
'
'Private Property Get ZS1S2_Ly(A As S1S2, W1%, W2%) As String()
'Dim S1$(), S2$()
'S1 = ZSplitLines(A.S1)
'S2 = ZSplitLines(A.S2)
'Dim M%, J%, O$(), Lin$, A1$, A2$, U1%, U2%
'    U1 = ZUB(S1)
'    U2 = ZUB(S2)
'    M = ZMax(U1, U2)
'Dim Spc1$, Spc2$
'    Spc1 = Space(W1)
'    Spc2 = Space(W2)
'For J = 0 To M
'   If J > U1 Then
'       A1 = Spc1
'   Else
'       A1 = ZStrAlignL(S1(J), W1)
'   End If
'   If J > U2 Then
'       A2 = Spc2
'   Else
'       A2 = ZStrAlignL(S2(J), W2)
'   End If
'   Lin = "| " + A1 + " | " + A2 + " |"
'   ZPush O, Lin
'Next
'ZS1S2_Ly = O
'End Property
'
'Private Sub ZS1S2_Push(O() As S1S2, M As S1S2)
'Dim N&
'N = ZS1S2_Sz(O)
'ReDim Preserve O(N)
'O(N) = M
'End Sub
'
'Private Property Get ZS1S2_Sz&(A() As S1S2)
'On Error Resume Next
'ZS1S2_Sz = UBound(A) + 1
'End Property
'
'Private Property Get ZS1S2_UB&(A() As S1S2)
'ZS1S2_UB = ZS1S2_Sz(A) - 1
'End Property
'
'Private Property Get ZSeed_Expand$(QVbl$, Ny$())
'Dim O$()
'Dim Sy$(): Sy = ZSplitVBar(QVbl)
'Dim J%, I
'For J = 0 To ZUB(Ny)
'    For Each I In Sy
'       ZPush O, Replace(I, "?", Ny(J))
'    Next
'Next
'ZSeed_Expand = ZJnCrLf(O)
'End Property
'
'Private Property Get ZSplitLines(A) As String()
'Dim B$: B = Replace(A, vbCrLf, vbLf)
'ZSplitLines = Split(B, vbLf)
'End Property
'
'Private Property Get ZSplitVBar(Vbl$) As String()
'ZSplitVBar = Split(Vbl, "|")
'End Property
'
'Private Property Get ZSrc(PjMdDotOrColonNm$) As String()
'ZSrc = ZMd_Src(ZMd(PjMdDotOrColonNm))
'End Property
'
'Private Property Get ZSrcLin_IsCd(A) As Boolean
'Dim L$: L = Trim(A)
'If A = "" Then Exit Property
'If Left(A, 1) = "'" Then Exit Property
'ZSrcLin_IsCd = True
'End Property
'
'Private Property Get ZSrcLin_IsMth(A) As Boolean
'ZSrcLin_IsMth = ZIs_FunTy(ZLin_T1(ZSrcLin_RmvMdy(A)))
'End Property
'
'Private Property Get ZSrcLin_MthNm$(A)
'Dim B$: B = ZLin_ShiftMthTy(ZSrcLin_RmvMdy(A))
'If A = "" Then Exit Property
'ZSrcLin_MthNm = ZLin_Nm(A)
'End Property
'
'Private Property Get ZSrcLin_RmvMdy$(L)
'Dim A$
'A = "Public ": If ZIs_Pfx(L, A) Then ZSrcLin_RmvMdy = ZRmvPfx(L, A): Exit Property
'A = "Friend ": If ZIs_Pfx(L, A) Then ZSrcLin_RmvMdy = ZRmvPfx(L, A): Exit Property
'A = "Private ": If ZIs_Pfx(L, A) Then ZSrcLin_RmvMdy = ZRmvPfx(L, A): Exit Property
'ZSrcLin_RmvMdy = L
'End Property
'
'Private Property Get ZSrc_MthBdyLines$(A$(), MthLx)
'Dim P1$
'    P1 = ZSrc_MthRmkLines(A, MthLx)
'Dim P2$
'    Dim L2%
'    L2 = ZSrc_EndLx(A, MthLx)
'    P2 = Join(ZAyWhFmTo(A, MthLx, L2), vbCrLf)
'If P1 = "" Then
'    ZSrc_MthBdyLines = P2
'Else
'    ZSrc_MthBdyLines = P1 & vbCrLf & P2
'End If
'End Property
'
'Private Property Get ZSrc_EndLx(A$(), MthLx)
'Dim F$: F = "End " & ZLin_FunTy(A(MthLx))
'Dim J%
'For J = MthLx + 1 To ZUB(A)
'    If ZIs_Pfx(A(J), F) Then ZSrc_EndLx = J: Exit Property
'Next
'End Property
'
'Private Property Get ZSrc_MthRmkLines$(A$(), MthLx)
'Dim O$(), J%, L$, I%
'Dim Lx&: Lx = ZSrc_MthRmkLx(A, MthLx)
'
'For J = Lx To MthLx - 1
'    L = Trim(A(J))
'    If L = "" Or L = "'" Then
'    ElseIf Left(L, 1) = "'" Then
'        ZPush O, L
'    Else
'        ' Er in ZSrc_MthRmkLx
'        Stop
'    End If
'Next
'ZSrc_MthRmkLines = Join(O, vbCrLf)
'End Property
'
'Private Property Get ZSrc_MthRmkLx&(A$(), MthLx)
'Dim M1&
'    Dim J&
'    For J = MthLx - 1 To 0 Step -1
'        If ZSrcLin_IsCd(A(J)) Then
'            M1 = J
'            GoTo M1IsFnd
'        End If
'    Next
'    M1 = -1
'M1IsFnd:
'Dim M2&
'    For J = M1 + 1 To MthLx - 1
'        If Trim(A(J)) <> "" Then
'            M2 = J
'            GoTo M2IsFnd
'        End If
'    Next
'    M2 = MthLx
'M2IsFnd:
'ZSrc_MthRmkLx = M2
'End Property
'
'Private Property Get ZSrc_DclLinCnt%(A$())
'Dim I&
'    I = ZSrc_FstMthLx(A)
'    If I = -1 Then
'        ZSrc_DclLinCnt = ZSz(A)
'        Exit Property
'    End If
'    I = ZSrc_MthRmkLx(A, I)
'Dim O&, L$
'    For I = I - 1 To 0 Step -1
'        If ZSrcLin_IsCd(A(I)) Then
'            O = I + 1
'            GoTo X
'        End If
'    Next
'X:
'ZSrc_DclLinCnt = O
'End Property
'
'Private Property Get ZSrc_DclLines$(A$())
'ZSrc_DclLines = Join(ZSrc_DclLy(A), vbCrLf)
'End Property
'
'Private Property Get ZSrc_DclLy(A$()) As String()
'If ZSz(A) = 0 Then Exit Property
'Dim N&
'   N = ZSrc_DclLinCnt(A)
'If N <= 0 Then Exit Property
'ZSrc_DclLy = ZLy_TrimEnd(ZAyFstNEle(A, N))
'End Property
'
'Private Property Get ZSrc_Dic(A$(), PjNm$, MdNm$) As Dictionary
'Dim O As Dictionary:
'If ZSz(A) = 0 Then
'    Set O = New Dictionary
'    O.Add ZFmtQQ("?:?:*Empty Md", PjNm, MdNm), ""
'    Set ZSrc_Dic = O
'    Exit Property
'End If
'Dim B() As S1S2: B = ZSrc_MthS1S2Ay(A, PjNm, MdNm)
'Set O = ZS1S2Ay_Dic(B)
'Dim D$: D = ZSrc_DclLines(A)
'    If D <> "" Then O.Add ZFmtQQ("?:?:*Dcl", PjNm, MdNm), D
'
'Set ZSrc_Dic = O
'End Property
'
'Private Property Get ZSrc_FstMthLx&(A$())
'Dim J%
'For J = 0 To ZUB(A)
'   If ZSrcLin_IsMth(A(J)) Then
'       ZSrc_FstMthLx = J
'       Exit Property
'   End If
'Next
'ZSrc_FstMthLx = -1
'End Property
'
'Private Property Get ZSrc_MthKy(A$(), Optional PjNm$ = "Pj", Optional MdNm$ = "Md", Optional IsSngLinFmt As Boolean) As String()
'Dim A1&(): A1 = ZSrc_MthLxAy(A)
'If ZSz(A1) = 0 Then Exit Property
'Dim O$()
'    Dim K$
'    Dim MthLx
'    Dim L$
'    For Each MthLx In A1
'        ZPush O, ZMthLin_MthKey(A(MthLx), PjNm, MdNm, IsSngLinFmt)
'    Next
'ZSrc_MthKy = O
'End Property
'
'Private Property Get ZSrc_MthLxAy(A$()) As Long()
'If ZSz(A) = 0 Then Exit Function
'Dim O&(), I, J&
'   For Each I In A
'       If ZSrcLin_IsMth(I) Then ZPush O, J
'       J = J + 1
'   Next
'ZSrc_MthLxAy = O
'End Property
'
'Private Property Get ZSrc_MthNy(A$(), Optional MthNmPatn$ = ".") As String()
'Dim A1&(): A1 = ZSrc_MthLxAy(A)
'If ZSz(A1) = 0 Then Exit Property
'Dim O$()
'    Dim MthLx, L$, N$, R As RegExp
'    Set R = ZRe(MthNmPatn)
'    For Each MthLx In A1
'        L = A(MthLx)
'        N = ZMthLin_MthNm(L)
'        If ZReTst(R, N) Then
'            ZPushNoDup O, N
'        End If
'    Next
'ZSrc_MthNy = ZAySrt(O)
'End Property
'
'Private Property Get ZSrc_MthS1S2Ay(A$(), PjNm$, MdNm$) As S1S2()
'Dim A1&(): A1 = ZSrc_MthLxAy(A)
'If ZSz(A1) = 0 Then Exit Property
'Dim O() As S1S2
'    Dim K$
'    Dim MthLx
'    Dim L$
'    For Each MthLx In A1
'        K = ZMthLin_MthKey(A(MthLx), PjNm, MdNm)
'        L = ZSrc_MthBdyLines(A, MthLx)
'        ZS1S2_Push O, ZS1S2(K, L)
'    Next
'ZSrc_MthS1S2Ay = O
'End Property
'
'Private Property Get ZSrc_SrtRpt(A$(), PjNm$, MdNm$) As DCRslt
'Dim B$(): B = ZSrc_SrtedLy(A)
'Dim A1 As Dictionary
'Dim B1 As Dictionary
'Set A1 = ZSrc_Dic(A, PjNm, MdNm)
'Set B1 = ZSrc_Dic(B, PjNm, MdNm)
'ZSrc_SrtRpt = ZDic_Cmp(A1, B1)
'End Property
'
'Private Property Get ZSrc_SrtRptLy(A$(), PjNm$, MdNm$) As String()
'ZSrc_SrtRptLy = ZDCRslt_Ly(ZSrc_SrtRpt(A, PjNm, MdNm))
'End Property
'
'Private Property Get ZSrc_SrtedBdyLines$(A$())
'If ZSz(A) = 0 Then Exit Property
'Dim B() As S1S2
'   B = ZSrc_MthS1S2Ay(A, "", "")
'Dim I&()
'   I = ZAySrtIntoIxAy(ZS1S2Ay_Sy1(B))
'Dim O$()
'Dim J%
'   For J = 0 To ZUB(I)
'       ZPush O, vbCrLf & B(I(J)).S2
'   Next
'ZSrc_SrtedBdyLines = Join(O, vbCrLf)
'End Property
'
'Private Property Get ZSrc_SrtedLines$(A$())
'Dim O$(), A1$, A2$, A3$, A4$
'A1 = ZSrc_DclLines(A)
'A2 = ZLines_TrimEnd(ZSrc_DclLines(A))
'A3 = ZSrc_SrtedBdyLines(A)
'A4 = ZLasChr(A3)
'If A4 = vbCr Or A4 = vbLf Then Stop
'ZPushNonEmp O, A2
'ZPushNonEmp O, A3
'ZSrc_SrtedLines = Join(O, vbCrLf)
'End Property
'
'Private Property Get ZSrc_SrtedLy(A$()) As String()
'ZSrc_SrtedLy = ZSplitLines(ZSrc_SrtedLines(A))
'End Property
'
'Private Property Get ZSsl_Sy(Ssl) As String()
'ZSsl_Sy = Split(Trim(ZRpl_DblSpc(Ssl)), " ")
'End Property
'
'Private Property Get ZStrAlignL$(S$, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
'Const CSub$ = "ZStrAlignL"
'Dim L%: L = Len(S)
'If L > W Then
'    If ErIfNotEnoughWdt Then
'        Stop
'        'Er CSub, "Len({S)) > {W}", S, W
'    End If
'    If DoNotCut Then
'        ZStrAlignL = S
'        Exit Property
'    End If
'End If
'
'If W >= L Then
'    ZStrAlignL = S & Space(W - L)
'    Exit Property
'End If
'If W > 2 Then
'    ZStrAlignL = Left(S, W - 2) + ".."
'    Exit Property
'End If
'ZStrAlignL = Left(S, W)
'End Property
'
'Private Sub ZStr_Brw(A$)
'Dim T$:
'T = ZTmpFt
'ZStr_Wrt A, T
'Shell ZFmtQQ("code.cmd ""?""", T), vbMaximizedFocus
''Shell ZFmtQQ("notepad.exe ""?""", T), vbMaximizedFocus
'End Sub
'
'Private Property Get ZStr_Dup$(S, N%)
'Dim O$, J%
'For J = 0 To N - 1
'    O = O & S
'Next
'ZStr_Dup = O
'End Property
'
'Private Property Get ZStr_Ny(A) As String()
'Dim O$: O = ZRpl_Pun(A)
'Dim O1$(): O1 = ZAyUniqAy(ZSsl_Sy(O))
'Dim O2$()
'Dim J%
'For J = 0 To ZUB(O1)
'    If Not ZIs_Digit(ZFstChr(O1(J))) Then ZPush O2, O1(J)
'Next
'ZStr_Ny = O2
'End Property
'
'Private Sub ZStr_Wrt(A, Ft$, Optional IsNotOvrWrt As Boolean)
'ZFso.CreateTextFile(Ft, Overwrite:=Not IsNotOvrWrt).Write A
'End Sub
'
'Private Property Get ZSz&(Ay)
'On Error Resume Next
'ZSz = UBound(Ay) + 1
'End Property
'
'Private Property Get ZTmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
'Dim Fnn$
'If Fnn0 = "" Then
'    Fnn = ZTmpNm
'Else
'    Fnn = Fnn0
'End If
'ZTmpFfn = ZTmpPth(Fdr) & Fnn & Ext
'End Property
'
'Private Property Get ZTmpFt$(Optional Fdr$, Optional Fnn$)
'ZTmpFt = ZTmpFfn(".txt", Fdr, Fnn)
'End Property
'
'Private Property Get ZTmpNm$()
'Static X&
'ZTmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
'X = X + 1
'End Property
'
'Private Property Get ZTmpPth$(Optional Fdr$)
'Dim X$
'   If Fdr <> "" Then
'       X = Fdr & "\"
'   End If
'Dim O$
'   O = ZTmpPthHom & X:   ZPth_Ens O
'   O = O & ZTmpNm & "\": ZPth_Ens O
'   ZPth_Ens O
'ZTmpPth = O
'End Property
'
'Private Property Get ZTmpPthHom$()
'Static X$
'If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
'ZTmpPthHom = X
'End Property
'
'Private Property Get ZToStr$(A)
'If ZIs_Prim(A) Then ZToStr = A: Exit Property
'If ZIs_Nothing(A) Then ZToStr = "#Nothing": Exit Property
'If IsEmpty(A) Then ZToStr = "#Empty": Exit Property
'If IsObject(A) Then
'    Dim T$
'    T = TypeName(A)
'    Select Case T
'    Case "CodeModule"
'        Dim M As CodeModule
'        Set M = A
'        ZToStr = ZFmtQQ("*Md{?}", M.Parent.Name)
'        Exit Property
'    End Select
'    ZToStr = "*" & T
'    Exit Property
'End If
'
'If IsArray(A) Then
'    Dim Ay: Ay = A: ReDim Ay(0)
'    T = TypeName(Ay(0))
'    ZToStr = "*[" & T & "]"
'    Exit Property
'End If
'Stop
'End Property
'
'Private Property Get ZUB&(Ay)
'ZUB = ZSz(Ay) - 1
'End Property
'
'Private Sub ZVbe_Export(A As VBE)
'ZOyDo ZVbe_PjAy(A), "ZPj_Export"
'End Sub
'
'Private Property Get ZVbe_MthNy(A As VBE, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".") As String()
'Dim Ay() As VBProject: Ay = ZVbe_PjAy(A)
'If ZSz(Ay) = 0 Then Exit Property
'Dim I, O$()
'For Each I In Ay
'    ZPushAy O, ZPj_MthNy(ZCvPj(I), MthNmPatn, MdNmPatn)
'Next
'ZVbe_MthNy = O
'End Property
'
'Private Property Get ZVbe_PjAy(A As VBE) As VBProject()
'Dim I, O() As VBProject
'For Each I In A.VBProjects
'    ZPushObj O, I
'Next
'ZVbe_PjAy = O
'End Property
'
'Private Property Get ZVbe_PjNy(A As VBE) As String()
'ZVbe_PjNy = ZItr_Ny(A.VBProjects)
'End Property
'
'Private Property Get ZVbe_SrtRptLy(A As VBE) As String()
'Dim Ay() As VBProject: Ay = ZVbe_PjAy(A)
'Dim O$(), I, M As VBProject
'For Each I In Ay
'    Set M = I
'    ZPushAy O, ZPj_SrtRptLy(M)
'Next
'ZVbe_SrtRptLy = O
'End Property
'
'Private Property Get ZWb_AddWs(A As Workbook, Optional WsNm$ = "Sheet1") As Worksheet
'Dim O As Worksheet
'Set O = A.Sheets.Add(A.Sheets(1))
'If WsNm <> "" Then
'   O.Name = WsNm
'End If
'Set ZWb_AddWs = O
'End Property
'
'Private Function ZXls() As Excel.Application
'Static Y As Excel.Application
'On Error GoTo X
'Dim A$: A = Y.Name
'Set ZXls = Y
'Exit Function
'X:
'Set Y = New Excel.Application
'Set ZXls = Y
'End Function
'
'Private Property Get ZZMd() As CodeModule
'Set ZZMd = ZCurVbe.VBProjects("QVb").VBComponents("M_A").CodeModule
'End Property
'
'Private Property Get ZZSrc() As String()
'ZZSrc = ZMd_Src(ZMd("M_Tool"))
'End Property
'
'Private Sub ZZ_Dcl_BefAndAft_Srt()
'Const MdNm$ = "VbStrRe"
'Dim A$() ' Src
'Dim B$() ' Src->Srt
'Dim A1$ 'Src->Dcl
'Dim B1$ 'Src->Src->Dcl
'A = ZMd_Src(ZMd("QSqTp.SalRpt"))
'B = ZSrc_SrtedLy(A)
'A1 = ZSrc_DclLines(A)
'B1 = ZSrc_DclLines(B)
'If A1 <> B1 Then Stop
'End Sub
'
'Private Sub ZZ_Go_Mth()
'Go_Mth "QTool.M_Tool.ZDotDotNm_BrkAsg"
'End Sub
'
'Private Sub ZZ_PjSrtRptWb()
'Dim O As Workbook: Set O = ZPj_SrtRptWb(ZCurPj, Vis:=True)
'Stop
'End Sub
'
'Private Sub ZZ_Pj_Compile()
'ZPj_Compile ZPj("QVb")
'End Sub
'
'Private Sub ZZ_ReMatch()
'Dim A As MatchCollection
'Dim R  As RegExp: Set R = ZRe("m[ae]n")
'Set A = ZReMatch(R, "alskdflfmensdklf")
'Stop
'End Sub
'
'Private Sub ZZ_Shw_CurPj_SrtRptWb()
'Shw_CurPj_SrtRptWb ZCurPj
'End Sub
'
'Private Sub ZZ_ZCurMdNm()
'Debug.Print ZCurMdNm
'End Sub
'
'Private Sub ZZ_ZCurVbe_PjNy()
'ZAyDmp ZCurVbe_PjNy
'End Sub
'
'Private Sub ZZ_ZMd_Gen_TstSub()
'ZMd_Gen_TstSub ZZMd
'End Sub
'
'Private Sub ZZ_ZMd_MthNy()
'ZAyBrw ZMd_MthNy(ZCurMd)
'End Sub
'
'Private Sub ZZ_ZMd_Rmv_TstSub()
'ZMd_Rmv_TstSub ZZMd
'End Sub
'
'Private Sub ZZ_ZMd_SrtedLines()
'ZStr_Brw ZMd_SrtedLines(ZMd("QVb.M_Ay"))
'End Sub
'
'Private Sub ZZ_ZMd_TstSub_BdyLines()
'Debug.Print ZMd_TstSub_BdyLines(ZZMd)
'End Sub
'
'Private Sub ZZ_ZMd_TstSub_Lno()
'Debug.Print ZMd_TstSub_Lno(ZZMd)
'End Sub
'
'Private Sub ZZ_ZPj()
'ZAss "QAcs" = ZPj("QAcs").Name
'End Sub
'
'Private Sub ZZ_ZPj_MthS1S2Ay()
'Dim A() As S1S2: A = ZPj_MthS1S2Ay(ZPj("QVb"))
'ZAyBrw ZS1S2Ay_FmtLy(A)
'End Sub
'
'Private Sub ZZ_ZPj_RfLy()
'ZAyBrw ZPj_RfLy(ZCurPj)
'End Sub
'
'Private Sub ZZ_ZPj_SrtRptLy()
'ZAyBrw ZPj_SrtRptLy(ZPj("QSqTp"))
'End Sub
'
'Private Sub ZZ_ZPj_TstClass_Bdy()
'Debug.Print ZPj_TstClass_Bdy(ZPj("QVb"))
'End Sub
'
'Private Sub ZZ_ZReRpl()
'Dim R As RegExp: Set R = ZRe("(.+)(m[ae]n)(.+)")
'Dim Act$: Act = ZReRpl(R, "a men is male", "$1male$3")
'ZAss Act = "a male is male"
'End Sub
'
'Private Sub ZZ_ZS1S2Ay_FmtLy()
'Dim Act$()
'Dim A() As S1S2
'ReDim A(4)
'Dim A1$, A2$
'Dim I%
'I = 0: A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df": GoSub XX
'I = 1: A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub XX
'I = 2: A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df": GoSub XX
'I = 3: A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df": GoSub XX
'I = 4: A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df": GoSub XX
'
'Act = ZS1S2Ay_FmtLy(A)
'ZAyBrw Act
'Exit Sub
'XX:
'    A(I) = ZS1S2(ZRpl_VBar(A1), ZRpl_VBar(A2))
'    Return
'End Sub
'
'Private Sub ZZ_ZSrc_DclLinCnt()
'Dim B$(), A%
''
'B = ZZSrc
'A = ZSrc_DclLinCnt(B)
'ZAss A = 43
''
'B = ZMd_Src(ZMd("QSqTp.SqTp2"))
'A = ZSrc_DclLinCnt(B)
'ZAss A = 688
'End Sub
'
'Private Sub ZZ_ZSrc_DclLines()
'Const P$ = "QSqTp"
'Const M$ = "SalRpt__CrdTyLvs_CrdExpr__Tst"
'Dim Md As CodeModule: Set Md = ZCurVbe.VBProjects(P).VBComponents(M).CodeModule
'Dim A$(): A = ZMd_Src(Md)
'Stop
'Dim B$: B = ZSrc_DclLines(A)
'Stop
'ZStr_Brw B
'End Sub
'
'Private Sub ZZ_ZSrc_MthS1S2Ay()
'Dim A() As S1S2: A = ZSrc_MthS1S2Ay(ZSrc("QVb.M_Ay"), "QTool", "M_Ay")
'ZAyBrw ZS1S2Ay_FmtLy(A)
'End Sub
'
'Private Sub ZZ_ZSrc_SrtRptLy()
'ZAyBrw ZSrc_SrtRptLy(ZZSrc, "QTool", "M_Tool")
'End Sub
'
'Private Sub ZZ_ZSrc_SrtedBdyLines()
'ZStr_Brw ZSrc_SrtedBdyLines(ZZSrc)
'End Sub
'
'Private Sub ZZ_ZSrc_SrtedLines()
'ZStr_Brw ZSrc_SrtedLines(ZZSrc)
'End Sub
'
'Private Sub ZZ_ZSrc_SrtedLy()
'ZAyBrw ZSrc_SrtedLy(ZZSrc)
'End Sub
'
'Private Sub ZZ_ZStr_Ny()
'Dim S$: S = ZMd_Lines(ZCurMd)
'ZAyBrw ZAySrt(ZStr_Ny(S))
'End Sub
'
'Private Sub ZZ_ZVbe_MthNy()
'ZAyBrw ZVbe_MthNy(ZCurVbe)
'End Sub
'
'Private Property Get ZCurMthNm$()
'
'End Property
'
'Sub Rmk_Mth_Bdy()
'ZMdMth_Rmk_Bdy ZCurMd, ZCurMthNm
'End Sub
'
'Private Property Get ZSrc_MthToLno%(A$(), FmLno%)
'Stop '
'ZSrc_MthToLno = 0
'End Property
'
'Private Property Get ZSrc_MthFmLno%(A$(), MthNm$)
'Dim O%, I
'For Each I In A
'    O = O + 1
'    If ZSrcLin_MthNm(I) = MthNm Then
'        ZSrc_MthFmLno = O
'        Exit Property
'    End If
'Next
'ZSrc_MthFmLno = O
'End Property
'
'Private Property Get ZSrc_MthFmToLno(A$(), MthNm$) As FmToLno
'If ZSz(A) = 0 Then Exit Property
'Dim F%, T%
'F = ZSrc_MthFmLno(A, MthNm)
'T = ZSrc_MthToLno(A, F)
'With ZSrc_MthFmToLno
'    .FmLno = T
'    .ToLno = T
'End With
'End Property
'
'Private Property Get ZMdMth_BdyFmToLno(A As CodeModule, MthNm$) As FmToLno
'ZMdMth_BdyFmToLno = ZSrc_MthBdyFmToLno(ZMd_Src(A), MthNm)
'End Property
'
'Private Property Get ZSrc_MthBdyFmToLno(A$(), MthNm$) As FmToLno
'Dim P As FmToLno
'    P = ZSrc_MthFmToLno(A, MthNm)
'Dim FmLno%, Fnd As Boolean
'For FmLno = P.FmLno To P.ToLno
'    If Not ZLasChr(A(FmLno)) = "_" Then
'        FmLno = FmLno + 1
'        Fnd = True
'        Exit For
'    End If
'Next
'If Not Fnd Then Stop
'With ZSrc_MthBdyFmToLno
'    .FmLno = FmLno
'    .ToLno = P.ToLno - 1
'End With
'End Property
'
'Private Sub ZMdMth_Rmk_Bdy(A As CodeModule, MthNm$)
'Dim P As FmToLno
'    P = ZMdMth_BdyFmToLno(A, MthNm)
'Dim J%, L$
'For J = P.FmLno To P.ToLno
'    L = A.Lines(J, 1)
'    A.ReplaceLine J, "'" & L
'Next
'A.InsertLines P.FmLno, "Stop" & " '"
'End Sub
'
'Private Sub ZMd_Cpy_ToPj(A As CodeModule, ToPj As VBProject)
'Dim MdNm$
'Dim FmPj As VBProject
'    Set FmPj = ZMd_Pj(A)
'    MdNm = A.Parent.Name
'If MdNm = "M_Tool" And ZCurPj = "QTool" Then
'    Debug.Print "Md(QTool.M_Tool) cannot be moved"
'    Exit Sub
'End If
'If ZPj_HasMdNm(ToPj, MdNm) Then
'    Debug.Print ZFmtQQ("ZMd_Cpy_ToPj: Md(?) exists in TarPj(?).  Skip moving", MdNm, ToPj.Name)
'    Exit Sub
'End If
'Dim ZTmpFil$
'    ZTmpFil = ZTmpFfn(".txt")
'    Dim SrcCmp As VBComponent
'    Set SrcCmp = A.Parent
'    SrcCmp.Export ZTmpFil
'    If SrcCmp.Type = vbext_ct_ClassModule Then
'        ZFt_RmvFst4Lines ZTmpFil
'    End If
'Dim TarCmp As VBComponent
'    Set TarCmp = ToPj.VBComponents.Add(A.Parent.Type)
'    TarCmp.CodeModule.AddFromFile ZTmpFil
'Kill ZTmpFil
'ZPj_Sav ToPj
'Debug.Print ZFmtQQ("ZMd_Cpy_ToPj: Md(?) is moved from SrcPj(?) to TarPj(?).", MdNm, FmPj.Name, ToPj.Name)
'End Sub
'Private Property Get ZFxaNm_Fxa$(A)
'Stop '
'End Property
'
'Private Sub ZPj_Crt_Fxa(A As VBProject, FxaNm$)
'Dim F$
'F = ZFxaNm_Fxa(FxaNm)
'End Sub
'Private Sub ZMd_Dlt(A As CodeModule)
'Dim M$, P$, Pj As VBProject
'    M = ZMd_Nm(A)
'    Set Pj = ZMd_Pj(A)
'    P = Pj.Name
'A.Parent.Collection.Remove A.Parent
'ZPj_Sav Pj
'Debug.Print ZFmtQQ("ZMd_Dlt: Md(?) is deleted from Pj(?)", M, P)
'End Sub
'Sub Cpy_Mbr(FmPjMbrDotNm$)
'ZMd_Cpy_ToPj ZMd(FmPjMbrDotNm), ZCurPj
'End Sub
