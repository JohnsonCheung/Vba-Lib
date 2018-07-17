Attribute VB_Name = "G_Tool1"
'Option Explicit
'Type Either
'    IsLeft As Boolean
'    Left As Variant
'    Right As Variant
'End Type
'Type FmToLno
'    FmLno As Integer
'    ToLno As Integer
'End Type
'Type DCRslt
'    Nm1 As String
'    Nm2 As String
'    AExcess As New Dictionary
'    BExcess As New Dictionary
'    ADif As New Dictionary
'    BDif As New Dictionary
'    Sam As New Dictionary
'End Type
'Type DicPair
'    A As Dictionary
'    B  As Dictionary
'End Type
'Type SyPair
'    Sy1() As String
'    Sy2() As String
'End Type
'Type MdSrtRpt
'    MdNy() As String
'    RptDic As Dictionary ' K is Module Name, V is DicCmpRsltLy
'End Type
'Type LCC
'    Lno As Integer
'    C1 As Integer
'    C2 As Integer
'End Type
'Type LCCOpt
'    Som As Boolean
'    LCC As LCC
'End Type
'Property Get ZZA()
'End Property
'Property Let ZZA(A)
'End Property
'
'Function MthNm_MthPfx$(A)
'Dim P%: P = InStr(A, "_")
'If P > 0 Then MthNm_MthPfx = Left(A, P): Exit Function
'Dim C$
'Dim Q%
'For P = 1 To Len(A)
'    C = Asc(Mid(A, P, 1))
'    If 97 <= C And C <= 122 Then
'        For Q = P + 1 To Len(A)
'            C = Asc(Mid(A, Q, 1))
'            If 65 <= C And C <= 90 Then
'                MthNm_MthPfx = Left(A, Q - 1)
'                Exit Function
'            End If
'        Next
'        MthNm_MthPfx = A
'        Exit Function
'    End If
'Next
'MthNm_MthPfx = A
'Stop
'End Function
'Sub AyBrw(Ay)
'StrBrw Join(Ay, vbCrLf)
'End Sub
'
'Sub AyDmp(Ay)
'If Sz(Ay) = 0 Then Exit Sub
'Dim I
'For Each I In Ay
'    Debug.Print I
'Next
'End Sub
'
'Sub AyDo(Ay, DoMthNm$)
'If Sz(Ay) = 0 Then Exit Sub
'Dim I
'For Each I In Ay
'    Run DoMthNm, I
'Next
'End Sub
'
'Sub AyWrt(Ay, Ft$)
'StrWrt JnCrLf(Ay), Ft
'End Sub
'
'
'
'Property Get AySrtInToIxAy__Ix&(Ix&(), A, V, Des As Boolean)
'Dim I, O&
'If Des Then
'    For Each I In Ix
'        If V > A(I) Then AySrtInToIxAy__Ix& = O: Exit Property
'        O = O + 1
'    Next
'    AySrtInToIxAy__Ix& = O
'    Exit Property
'End If
'For Each I In Ix
'    If V < A(I) Then AySrtInToIxAy__Ix& = O: Exit Property
'    O = O + 1
'Next
'AySrtInToIxAy__Ix& = O
'End Property
'
'Property Get AySrt__Ix&(Ay, V, Des As Boolean)
'Dim I, O&
'If Des Then
'    For Each I In Ay
'        If V > I Then AySrt__Ix = O: Exit Property
'        O = O + 1
'    Next
'    AySrt__Ix = O
'    Exit Property
'End If
'For Each I In Ay
'    If V < I Then AySrt__Ix = O: Exit Property
'    O = O + 1
'Next
'AySrt__Ix = O
'End Property
'
'Property Get DCRslt_Ly__AExcess(A As Dictionary) As S1S2()
'If A.Count = 0 Then Exit Property
'Dim O() As S1S2, K
'For Each K In A.Keys
'    PushObj O, S1S2(K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K), "!" & "Er AExcess")
'Next
'DCRslt_Ly__AExcess = O
'End Property
'
'Property Get DCRslt_Ly__BExcess(A As Dictionary) As S1S2()
'If A.Count = 0 Then Exit Property
'Dim O() As S1S2, K
'For Each K In A.Keys
'    PushObj O, S1S2("!" & "Er BExcess", K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K))
'Next
'DCRslt_Ly__BExcess = O
'End Property
'
'Property Get DCRslt_Ly__Dif(A As Dictionary, B As Dictionary) As S1S2()
'If A.Count <> B.Count Then Stop
'If A.Count = 0 Then Exit Property
'Dim O() As S1S2, K, S1$, S2$
'For Each K In A
'    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
'    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & B(K)
'    PushObj O, S1S2(S1, S2)
'Next
'DCRslt_Ly__Dif = O
'End Property
'
'Property Get DCRslt_Ly__Sam(A As Dictionary) As S1S2()
'If A.Count = 0 Then Exit Property
'Dim O() As S1S2, K
'For Each K In A.Keys
'    PushObj O, S1S2("*Same", K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K))
'Next
'DCRslt_Ly__Sam = O
'End Property
'
'Property Get PjMbrAy__X(A As VBProject, MbrNmPatn$, TyAy() As vbext_ComponentType) As CodeModule()
'Dim O() As CodeModule
'Dim Cmp As VBComponent
'Dim R As RegExp: If MbrNmPatn <> "." Then Set R = Re(MbrNmPatn)
'For Each Cmp In A.VBComponents
'    If AyHas(TyAy, Cmp.Type) Then
'        If MbrNmPatn = "." Then
'            PushObj O, Cmp.CodeModule
'        Else
'            If R.Test(Cmp.Name) Then
'                PushObj O, Cmp.CodeModule
'            End If
'        End If
'    End If
'Next
'PjMbrAy__X = O
'End Property
'
'Property Get AyAddPfx(Ay, Pfx) As String()
'If Sz(Ay) = 0 Then Exit Property
'Dim O$(), I
'For Each I In Ay
'    Push O, Pfx & I
'Next
'AyAddPfx = O
'End Property
'
'Property Get AyAddPfxSfx(A, P$, S$) As String()
'Dim I, O$()
'If Sz(A) = 0 Then Exit Property
'For Each I In A
'    Push O, P & I & S
'Next
'AyAddPfxSfx = O
'End Property
'
'Property Get AyAddSfx(Ay, Sfx) As String()
'If Sz(Ay) = 0 Then Exit Property
'Dim O$(), I
'For Each I In Ay
'    Push O, I & Sfx
'Next
'AyAddSfx = O
'End Property
'
'Property Get AyAlignL(Ay) As String()
'Dim W%: W = AyWdt(Ay) + 1
'If Sz(Ay) = 0 Then Exit Property
'Dim O$(), I
'For Each I In Ay
'    Push O, AlignL(I, W)
'Next
'AyAlignL = O
'End Property
'
'Property Get AyColl(Ay) As Collection
'Dim O As New Collection, I
'If Sz(Ay) = 0 Then Set AyColl = O: Exit Property
'For Each I In Ay
'    O.Add I
'Next
'Set AyColl = O
'End Property
'
'Property Get AyDblQuote(A) As String()
'Const C$ = """"
'AyDblQuote = AyAddPfxSfx(A, C, C)
'End Property
'
'Property Get AyFstNEle(A, N&)
'Dim O: O = A
'ReDim Preserve O(N - 1)
'AyFstNEle = O
'End Property
'
'Property Get AyHas(A, M) As Boolean
'Dim I: If Sz(A) = 0 Then Exit Property
'For Each I In A
'    If I = M Then AyHas = True: Exit Property
'Next
'End Property
'
'Property Get AyIns(A, Optional M, Optional At&)
'Dim N&: N = Sz(A)
'If 0 > At Or At > N Then
'    Stop
'End If
'Dim O
'    O = A
'    ReDim Preserve O(N)
'    Dim J&
'    For J = N To At + 1 Step -1
'        Asg O(J - 1), O(J)
'    Next
'    O(At) = M
'AyIns = O
'End Property
'
'Property Get AyIsAllEleEq(A) As Boolean
'If Sz(A) = 0 Then AyIsAllEleEq = True: Exit Property
'Dim J&
'For J = 1 To UB(A)
'    If A(0) <> A(J) Then Exit Property
'Next
'AyIsAllEleEq = True
'End Property
'
'Property Get AyLasEle(Ay)
'AyLasEle = Ay(UB(Ay))
'End Property
'
'Property Get AyMap(A, MapFunNm$)
'AyMap = AyMapInto(A, MapFunNm, EmpAy)
'End Property
'
'Property Get AyMapInto(A, MapFunNm$, OIntoAy)
'Erase OIntoAy
'Dim I
'If Sz(A) > 0 Then
'    For Each I In A
'        Push OIntoAy, Run(MapFunNm, I)
'    Next
'End If
'AyMapInto = OIntoAy
'End Property
'
'Property Get AyMapSy(A, MapFunNm$) As String()
'AyMapSy = AyMapInto(A, MapFunNm, EmpSy)
'End Property
'
'Property Get AyMinus(A, B)
'If Sz(B) = 0 Or Sz(A) = 0 Then AyMinus = A: Exit Property
'Dim O: O = A: Erase O
'Dim B1: B1 = B
'Dim V
'For Each V In A
'    If AyHas(B1, V) Then
'        B1 = AyRmvEle(B1, V)
'    Else
'        Push O, V
'    End If
'Next
'AyMinus = O
'End Property
'
'Property Get AyMinusAp(Ay, ParamArray AyAp())
'Dim O
'If Sz(Ay) = 0 Then O = Ay: Erase O: GoTo X
'O = Ay
'Dim Av(): Av = AyAp
'Dim Ay1, V
'For Each Ay1 In Av
'    O = AyMinus(O, Ay1)
'    If Sz(O) = 0 Then GoTo X
'Next
'X:
'AyMinusAp = O
'End Property
'
'Property Get AyPair_Dic(A1, A2) As Dictionary
'Dim N1&, N2&
'N1 = Sz(A1)
'N2 = Sz(A2)
'If N1 <> N2 Then Stop
'Dim O As New Dictionary
'Dim J&
'If Sz(A1) = 0 Then GoTo X
'For J = 0 To N1 - 1
'    O.Add A1(J), A2(J)
'Next
'X:
'Set AyPair_Dic = O
'End Property
'
'Function AyRgH(A, At As Range) As Range
'Set AyRgH = AtPutSq(At, AySqH(A))
'End Function
'
'Property Get AyRmvEle(Ay, M)
'Dim O, V: O = Ay: Erase O
'For Each V In Ay
'    If V <> M Then Push O, M
'Next
'AyRmvEle = O
'End Property
'
'Property Get AyRmvEmp(Ay)
'If Sz(Ay) = 0 Then AyRmvEmp = Ay: Exit Property
'Dim O: O = Ay: Erase O
'Dim I
'For Each I In Ay
'    If Not IsEmp(I) Then Push O, I
'Next
'AyRmvEmp = O
'End Property
'
'Property Get AySqH(A) As Variant()
'Dim O(), J&
'ReDim O(1 To 1, 1 To Sz(A))
'For J = 0 To UB(A)
'    O(1, J + 1) = A(J)
'Next
'AySqH = O
'End Property
'
'Property Get AySqV(Ay) As Variant()
'If Sz(Ay) = 0 Then Exit Property
'Dim O(), R&
'ReDim O(1 To Sz(Ay), 1 To 1)
'R = 0
'Dim V
'For Each V In Ay
'    R = R + 1
'    O(R, 1) = V
'Next
'AySqV = O
'End Property
'
'Property Get AySrt(Ay, Optional Des As Boolean)
'If Sz(Ay) = 0 Then AySrt = Ay: Exit Property
'Dim Ix&, V, J&
'Dim O: O = Ay: Erase O
'Push O, Ay(0)
'For J = 1 To UB(Ay)
'    O = AyIns(O, Ay(J), AySrt__Ix(O, Ay(J), Des))
'Next
'AySrt = O
'End Property
'
'Property Get AySrtIntoIxAy(Ay, Optional Des As Boolean) As Long()
'If Sz(Ay) = 0 Then Exit Property
'Dim Ix&, V, J&
'Dim O&():
'Push O, 0
'For J = 1 To UB(Ay)
'    O = AyIns(O, J, AySrtInToIxAy__Ix(O, Ay, Ay(J), Des))
'Next
'AySrtIntoIxAy = O
'End Property
'
'Property Get AyUniqAy(Ay)
'Dim O: O = Ay: Erase O
'If Sz(Ay) > 0 Then
'    Dim I
'    For Each I In Ay
'        PushNoDup O, I
'    Next
'End If
'AyUniqAy = O
'End Property
'
'Property Get AyWdt%(Ay)
'Dim W%, I: If Sz(Ay) = 0 Then Exit Property
'For Each I In Ay
'    W = Max(Len(I), W)
'Next
'AyWdt = W
'End Property
'
'Property Get AyWhFmTo(Ay, FmIx, ToIx)
'Dim O: O = Ay: Erase O
'Dim J&
'For J = FmIx To ToIx
'    Push O, Ay(J)
'Next
'AyWhFmTo = O
'End Property
'
'Sub MdAddFun(A As CodeModule, Nm$, IsFun As Boolean)
'Dim L$
'    Dim B$
'    B = IIf(IsFun, "Function", "Sub")
'    L = FmtQQ("? ?()|End ?", B, Nm, B)
'MdAppLines A, L
'MthGo Mth(A, Nm)
'End Sub
'
'Property Get AlignL$(A, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
'Const CSub$ = "AlignL"
'If ErIfNotEnoughWdt And DoNotCut Then
'    Stop
'    'Er CSub, "Both {ErIfNotEnoughWdt} and {DontCut} cannot be True", ErIfNotEnoughWdt, DoNotCut
'End If
'Dim S$: S = VarStr(A)
'AlignL = StrAlignL(S, W, ErIfNotEnoughWdt, DoNotCut)
'End Property
'
'Sub Asg(V, OV)
'If IsObject(V) Then
'   Set OV = V
'Else
'   OV = V
'End If
'End Sub
'
'Sub Ass(A As Boolean)
'If Not A Then Stop
'End Sub
'
'Property Get AtPutSq(A As Range, Sq) As Range
'Dim R As Range
'    Set R = AtReSz(A, Sq)
'R.Value = Sq
'Set AtPutSq = R
'End Property
'
'Property Get AtReSz(A As Range, Sq) As Range
'Set AtReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
'End Property
'
'Property Get Brk(A, Sep$, Optional IsNoTrim As Boolean) As S1S2
'Dim P&: P = InStr(A, Sep)
'If P = 0 Then Stop
'Dim S1$, S2$
'    S1 = Left(A, P - 1)
'    S2 = Mid(A, P + Len(Sep))
'If Not IsNoTrim Then
'    S1 = Trim(S1)
'    S2 = Trim(S2)
'End If
'Set Brk = S1S2(S1, S2)
'End Property
'
'Sub Brk2_Asg(A, Sep$, O1$, O2$)
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
'Property Get CmpTyAy_Of_Cls() As vbext_ComponentType()
'Dim T() As vbext_ComponentType
'T(0) = vbext_ct_ClassModule
'CmpTyAy_Of_Cls = T
'End Property
'
'Property Get CmpTyAy_Of_Cls_and_Md() As vbext_ComponentType()
'Dim T(1) As vbext_ComponentType
'T(0) = vbext_ct_ClassModule
'T(1) = vbext_ct_StdModule
'CmpTyAy_Of_Cls_and_Md = T
'End Property
'
'Property Get CmpTyAy_Of_Md() As vbext_ComponentType()
'Dim T(0) As vbext_ComponentType
'T(0) = vbext_ct_StdModule
'CmpTyAy_Of_Md = T
'End Property
'
'Property Get CmpTy_Nm$(A As vbext_ComponentType)
'Dim O$
'Select Case A
'Case vbext_ct_ClassModule: O = "*Cls"
'Case vbext_ct_StdModule: O = "*Md"
'Case Else: Stop
'End Select
'CmpTy_Nm = O
'End Property
'
'Sub CmpRmv(A As VBComponent)
'A.Collection.Remove A
'End Sub
'
'Property Get CollAddPfx(A As Collection, Pfx) As Collection
'Dim O As New Collection, I
'For Each I In A
'    O.Add Pfx & I
'Next
'Set CollAddPfx = O
'End Property
'
'Property Get CurCmp() As VBComponent
'Set CurCmp = CurMd.Parent
'End Property
'
'Property Get CurMd() As CodeModule
'Set CurMd = CurVbe.ActiveCodePane.CodeModule
'End Property
'
'Property Get CurMdNm$()
'CurMdNm = CurCmp.Name
'End Property
'
'Property Get CurMth() As Mth
'Dim Nm$: Nm = CurMthNm
'If Nm = "" Then Stop
'Set CurMth = Mth(CurMd, Nm)
'End Property
'Private Sub XX()
'Debug.Print "xx..."
'End Sub
'Property Get CurMthNm$()
'Dim L1&, L2&, C1&, C2&, K As vbext_ProcKind
'Dim O$
'With CurVbe.ActiveCodePane
'    On Error GoTo X
'    .GetSelection L1, C1, L2, C2
'    On Error GoTo 0
'    O = .CodeModule.ProcOfLine(L1, K)
'End With
'If O = "" Then Stop
'CurMthNm = O
'Exit Property
'X:
'End Property
'Property Get CurMdDNm$()
'CurMdDNm = MdDNm(CurMd)
'End Property
'Property Get CurMthDNm$()
'CurMthDNm = CurMdDNm & "." & CurMthNm
'End Property
'
'Property Get CurPj() As VBProject
'Set CurPj = CurVbe.ActiveVBProject
'End Property
'
'Property Get CurPjNm$()
'CurPjNm = CurPj.Name
'End Property
'Property Get MthDNm$(A As Mth)
'MthDNm = MdDNm(A.Md) & "." & A.Nm
'End Property
'Property Get MdDNm$(A As CodeModule)
'MdDNm = MdPjNm(A) & "." & MdNm(A)
'End Property
'
'Property Get CurVbe() As Vbe
'Set CurVbe = Excel.Application.Vbe
'End Property
'
'Property Get CurVbe_DupMdNy() As String()
'CurVbe_DupMdNy = VbeDupMdNy(CurVbe)
'End Property
'
'Property Get CurVbe_MdPjNy(MdNm$) As String()
'CurVbe_MdPjNy = VbeMdPjNy(CurVbe, MdNm)
'End Property
'
'Property Get CurVbe_MthNy(Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$) As String()
'CurVbe_MthNy = VbeMthNy(CurVbe, MthNmPatn, MdNmPatn, Mdy)
'End Property
'
'Property Get CurVbe_PjAy() As VBProject()
'CurVbe_PjAy = VbePjAy(CurVbe)
'End Property
'
'Property Get CurVbe_PjNy() As String()
'CurVbe_PjNy = VbePjNy(CurVbe)
'End Property
'
'Property Get CvMd(A) As CodeModule
'Set CvMd = A
'End Property
'
'Property Get CvPj(I) As VBProject
'Set CvPj = I
'End Property
'
'Property Get CvSy(A) As String()
'CvSy = A
'End Property
'
'Property Get DCRslt_IsSam(A As DCRslt) As Boolean
'With A
'If .ADif.Count > 0 Then Exit Property
'If .BDif.Count > 0 Then Exit Property
'If .AExcess.Count > 0 Then Exit Property
'If .BExcess.Count > 0 Then Exit Property
'End With
'DCRslt_IsSam = True
'End Property
'
'Property Get DCRslt_Ly(A As DCRslt) As String()
'With A
'Dim A1() As S1S2: A1 = DCRslt_Ly__AExcess(.AExcess)
'Dim A2() As S1S2: A2 = DCRslt_Ly__BExcess(.BExcess)
'Dim A3() As S1S2: A3 = DCRslt_Ly__Dif(.ADif, .BDif)
'Dim A4() As S1S2: A4 = DCRslt_Ly__Sam(.Sam)
'End With
'Dim O() As S1S2
'PushObj O, S1S2(A.Nm1, A.Nm2)
'O = S1S2Ay_Add(O, A1)
'O = S1S2Ay_Add(O, A2)
'O = S1S2Ay_Add(O, A3)
'O = S1S2Ay_Add(O, A4)
'DCRslt_Ly = S1S2Ay_FmtLy(O)
'End Property
'
'Property Get DftMdByMdNm(MdNm$) As CodeModule
'If MdNm = "" Then
'    Set DftMdByMdNm = CurMd
'Else
'    Set DftMdByMdNm = Md(MdNm)
'End If
'End Property
'
'Property Get DftNy(Ny0) As String()
'Dim T As VbVarType: T = VarType(Ny0)
'If T = vbEmpty Then Exit Property
'If IsMissing(Ny0) Then Exit Property
'If T = vbString Then
'    DftNy = SplitSsl(Ny0)
'    Exit Property
'End If
'DftNy = Ny0
'End Property
'
'Property Get DicPair_SamKeyDifValPair(A As Dictionary, B As Dictionary) As DicPair
'Dim K, A1 As New Dictionary, B1 As New Dictionary
'For Each K In A.Keys
'    If B.Exists(K) Then
'        If A(K) <> B(K) Then
'            A1.Add K, A(K)
'            B1.Add K, B(K)
'        End If
'    End If
'Next
'With DicPair_SamKeyDifValPair
'    Set .A = A1
'    Set .B = B1
'End With
'End Property
'
'Property Get DicAdd(A As Dictionary, B As Dictionary) As Dictionary
'Dim O  As New Dictionary, I
'For Each I In A.Keys
'    O.Add I, A(I)
'Next
'For Each I In B.Keys
'    O.Add I, B(I)
'Next
'Set DicAdd = O
'End Property
'
'Property Get DicClone(A As Dictionary) As Dictionary
'Dim O As New Dictionary, K
'For Each K In A.Keys
'    O.Add K, A(K)
'Next
'Set DicClone = O
'End Property
'
'Property Get DicCmp(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As DCRslt
'Dim O As DCRslt
'Set O.AExcess = DicMinus(A, B)
'Set O.BExcess = DicMinus(B, A)
'Set O.Sam = DicSam(A, B)
'With DicPair_SamKeyDifValPair(A, B)
'    Set O.ADif = .A
'    Set O.BDif = .B
'End With
'O.Nm1 = Nm1
'O.Nm2 = Nm2
'DicCmp = O
'End Property
'
'Property Get DicMinus(A As Dictionary, B As Dictionary) As Dictionary
'If A.Count = 0 Then Set DicMinus = New Dictionary: Exit Property
'If B.Count = 0 Then Set DicMinus = DicClone(A): Exit Property
'Dim O As New Dictionary, K
'For Each K In A.Keys
'   If Not B.Exists(K) Then O.Add K, A(K)
'Next
'Set DicMinus = O
'End Property
'
'Property Get DicSam(A As Dictionary, B As Dictionary) As Dictionary
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
'X: Set DicSam = O
'End Property
'
'Property Get DicWb(A As Dictionary, Optional Vis As Boolean) As Workbook
''Assume each dic keys is name and each value is lines
''Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
'Ass DicHasAllKeyIsNm(A)
'Ass DicHasAllValIsStr(A)
'Dim K, ThereIsSheet1 As Boolean
'Dim O As Workbook: Set O = NewWb
'Dim Ws As Worksheet
'For Each K In A.Keys
'    If K = "Sheet1" Then
'        Set Ws = O.Sheets("Sheet1")
'        ThereIsSheet1 = True
'    Else
'        Set Ws = O.Sheets.Add
'        Ws.Name = K
'    End If
'    Ws.Range("A1").Value = LinesSqV(A(K))
'Next
'X: Set Ws = O
'If Vis Then O.Application.Visible = True
'End Property
'
'Sub DDN_BrkAsg(A, O1$, O2$, O3$)
'Dim Ay$(): Ay = Split(A, ".")
'Select Case Sz(Ay)
'Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
'Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
'Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
'Case Else: Stop
'End Select
'End Sub
'
'Sub DrsBrw(A As Drs)
'Stop '
'End Sub
'
'Property Get DrsWs(A As Drs) As Worksheet
'Dim O As Worksheet, R As Range
'Set O = NewWs
'AyRgH A.Fny, WsA1(O)
'Set R = AtPutSq(WsRC(O, 2, 1), DrySq(A.Dry))
'Set DrsWs = O
'End Property
'
'Property Get DryNCol&(A())
'Dim O&, Dr
'For Each Dr In A
'    O = Max(O, Sz(Dr))
'Next
'DryNCol = O
'End Property
'
'Property Get DrySq(A() As Variant) As Variant()
'Dim NCol&, NRow&
'    NCol = DryNCol(A)
'    NRow = Sz(A)
'Dim O()
'ReDim O(1 To NRow, 1 To NCol)
'Dim C&, R&, Dr()
'    For R = 1 To NRow
'        Dr = A(R - 1)
'        For C = 1 To Min(Sz(Dr), NCol)
'            O(R, C) = Dr(C - 1)
'        Next
'    Next
'DrySq = O
'End Property
'
'Property Get DupMthFGpNy_Dr(A$()) As Variant()
'Dim Ny$(): Ny = A
'Stop '
'End Property
'
'Property Get DupMthFNyGp_Dry(Ny$()) As Variant()
''Given Ny: Each Nm in Ny is FunNm:PjNm.MdNm
''          It has at least 2 ele
''          Each FunNm is same
''Return: N-Dr of Fields {Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src}
''        where N = Sz(Ny)-1
''        where each-field-(*-1)-of-Dr comes from Ny(0)
''        where each-field-(*-2)-of-Dr comes from Ny(1..)
'
'Dim Md1$, Pj1$, Nm$
'    MthFNm_BrkAsg Ny(0), Nm, Pj1, Md1
'Dim Mth1 As New Mth
'    Mth1.Nm = Nm
'    Set Mth1.Md = Md(Pj1 & "." & Md1)
'Dim Src1$
'    Src1 = MthLines(Mth1)
'Dim Mdy1$, Ty1$
'    MthBrkAsg Mth1, Mdy1, Ty1
'Dim O()
'    Dim J%
'    For J = 1 To UB(Ny)
'        Dim Pj2$, Nm2$, Md2$
'            MthFNm_BrkAsg Ny(J), Nm2, Pj2, Md2: If Nm2 <> Nm Then Stop
'        Dim Mth2 As New Mth
'            Mth2.Nm = Nm
'            Set Mth2.Md = Md(Pj2 & "." & Md2)
'        Dim Src2$
'            Src2 = MthLines(Mth2)
'        Dim Mdy2$, Ty2$
'            MthBrkAsg Mth2, Mdy2, Ty2
'
'        Push O, Array(Nm, _
'                    Mdy1, Ty1, Pj1, Md1, _
'                    Mdy2, Ty2, Pj2, Md2, Src1, Src2, Pj1 = Pj2, Md1 = Md2, Src1 = Src2)
'    Next
'DupMthFNyGp_Dry = O
'End Property
'
'Property Get DupMthFNyGp_IsDup(Ny) As Boolean
'DupMthFNyGp_IsDup = AyIsAllEleEq(AyMap(Ny, "MthF_MthLines"))
'End Property
'
'Property Get DupMthFNy_GpAy(A$()) As Variant()
'Dim O(), J%, M$()
'Dim L$ ' LasMthNm
'L = Brk(A(0), ":").S1
'Push M, A(0)
'Dim B As S1S2
'For J = 1 To UB(A)
'    B = Brk(A(J), ":")
'    If L <> B.S1 Then
'        Push O, M
'        Erase M
'        L = B.S1
'    End If
'    Push M, A(J)
'Next
'If Sz(M) > 0 Then
'    Push O, M
'End If
'DupMthFNy_GpAy = O
'End Property
'
'Property Get DupMthFNy_SamMthBdyMthFNy(A$(), Vbe As Vbe) As String()
'Dim Gp(): Gp = DupMthFNy_GpAy(A)
'Dim O$(), N, Ny
'For Each Ny In Gp
'    If DupMthFNyGp_IsDup(Ny) Then
'        For Each N In Ny
'            Push O, N
'        Next
'    End If
'Next
'DupMthFNy_SamMthBdyMthFNy = O
'End Property
'
'Property Get DupFunDic_Add(DupFunDic As Dictionary, FunDic As Dictionary) As Dictionary
'
'End Property
'
'Property Get DupFunDic_Ly(A As Dictionary) As String()
'Stop '
'End Property
'
'Property Get EitherL(A) As Either
'Asg A, EitherL.Left
'EitherL.IsLeft = True
'End Property
'
'Property Get EitherR(A) As Either
'Asg A, EitherR.Right
'End Property
'
'Property Get EmpAy() As Variant()
'End Property
'
'Property Get EmpRfAy() As Reference()
'End Property
'
'Property Get EmpSy() As String()
'End Property
'
'Sub ErImposs()
'Stop ' Impossible
'End Sub
'
'Sub MthFNm_BrkAsg(A$, OFunNm$, OPjNm$, OMdNm$)
'With Brk(A, ":")
'    OFunNm = .S1
'    With Brk(.S2, ".")
'        OPjNm = .S1
'        OMdNm = .S2
'    End With
'End With
'End Sub
'
'Property Get MthFNm_Nm$(A$)
'MthFNm_Nm = Brk(A, ":").S1
'End Property
'
'Property Get MthFNy_DupMthFNy(A$(), Optional IsSamMthBdyOnly As Boolean) As String()
'If Sz(A) = 0 Then Exit Property
'Dim A1$(): A1 = AySrt(A)
'Dim O$(), M$(), J&, Nm$
'Dim L$ ' LasFunNm
'L = MthFNm_Nm(A1(0))
'Push M, A1(0)
'For J = 1 To UB(A1)
'    Nm = MthFNm_Nm(A1(J))
'    If L = Nm Then
'        Push M, A1(J)
'    Else
'        L = Nm
'        If Sz(M) = 1 Then
'            M(0) = A1(J)
'        Else
'            PushAy O, M
'            Erase M
'        End If
'    End If
'Next
'If Sz(M) > 1 Then
'    PushAy O, M
'End If
'MthFNy_DupMthFNy = O
'End Property
'
'Property Get MthFNm_Mth(A) As Mth
'Set MthFNm_Mth = MthDNm_Mth(MthFNm_MthDNm(A))
'End Property
'
'Function MthFNm_MthLines$(A)
'MthFNm_MthLines = MthLines(MthFNm_Mth(A))
'End Function
'
'Property Get MthFNm_MthDNm$(A)
'With Brk(A, ":")
'    MthFNm_MthDNm = .S2 & "." & .S1
'End With
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
'Sub FfnCpyToPth(A, ToPth$, Optional OvrWrt As Boolean)
'Fso.CopyFile A, ToPth$ & FfnFn(A), OvrWrt
'End Sub
'
'Sub FfnDlt(A)
'On Error Resume Next
'Kill A
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
'Property Get FfnFn$(A)
'Dim P%: P = InStrRev(A, "\")
'If P = 0 Then FfnFn = A: Exit Property
'FfnFn = Mid(A, P + 1)
'End Property
'
'Property Get FfnFnn$(A)
'FfnFnn = FfnRmvExt(FfnFn(A))
'End Property
'
'Property Get FfnIsExist(A) As Boolean
'FfnIsExist = Fso.FileExists(A)
'End Property
'
'Property Get FfnPth$(A)
'Dim P%: P = InStrRev(A, "\")
'If P = 0 Then Exit Property
'FfnPth = Left(A, P)
'End Property
'
'Property Get FfnRmvExt$(A)
'Dim P%: P = InStrRev(A, ".")
'If P = 0 Then FfnRmvExt = Left(A, P): Exit Property
'FfnRmvExt = Left(A, P - 1)
'End Property
'
'Property Get FmtQQ$(QQVbl$, ParamArray Ap())
'Dim O$: O = Replace(QQVbl, "|", vbCrLf)
'Dim Av(): Av = Ap
'Dim I
'For Each I In Av
'    O = Replace(O, "?", I, Count:=1)
'Next
'FmtQQ = O
'End Property
'
'Property Get Fso() As FileSystemObject
'Set Fso = New FileSystemObject
'End Property
'
'Property Get FstChr$(A)
'FstChr = Left(A, 1)
'End Property
'
'Sub FtRmvFst4Lines(Ft$)
'Dim A$: A = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
'Dim B$: B = Left(A, 55)
'Dim C$: C = Mid(A, 56)
'Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
'If B <> B1 Then Stop
'Fso.CreateTextFile(Ft, True).Write C
'End Sub
'
'Sub FxaNm_Crt(A)
'Fxa_Crt FxaNm_Fxa(A)
'End Sub
'
'Property Get FxaNm_Fxa$(A)
'FxaNm_Fxa = PjPth(CurPj) & A & ".xlam"
'End Property
'
'Sub Fxa_Crt(A)
'If FfnIsExist(A) Then Stop: Exit Sub
'Dim X As Excel.Application
'Set X = Xls
'If XlsHasAddInFn(X, FfnFn(A)) Then Stop: Exit Sub
'Dim O As Workbook
'Set O = X.Workbooks.Add
'O.SaveAs A, XlFileFormat.xlOpenXMLAddIn
'X.AddIns.Add(A).Installed = True
'O.Close
'End Sub
'
'Property Get HasPfx(S, Pfx$) As Boolean
'HasPfx = Left(S, Len(Pfx)) = Pfx
'End Property
'
'Property Get HasSubStr(A, SubStr$) As Boolean
'HasSubStr = InStr(A, SubStr) > 0
'End Property
'
'Property Get WdtAy_HdrLin$(A%())
'Dim O$(), W
'For Each W In A
'    Push O, StrDup("-", W + 2)
'Next
'WdtAy_HdrLin = "|" + Join(O, "|") + "|"
'End Property
'
'Property Get DicHasAllKeyIsNm(A As Dictionary) As Boolean
'Dim K
'For Each K In A.Keys
'    If Not IsNm(K) Then Exit Property
'Next
'DicHasAllKeyIsNm = True
'End Property
'
'Property Get DicHasAllValIsStr(A As Dictionary) As Boolean
'Dim K
'For Each K In A.Keys
'    If Not IsStr(A(K)) Then Exit Property
'Next
'DicHasAllValIsStr = True
'End Property
'
'Property Get IsDigit(A) As Boolean
'IsDigit = "0" <= A And A <= "9"
'End Property
'
'Property Get IsEmp(V) As Boolean
'IsEmp = True
'If IsMissing(V) Then Exit Property
'If IsNothing(V) Then Exit Property
'If IsEmpty(V) Then Exit Property
'If IsStr(V) Then
'   If V = "" Then Exit Property
'End If
'If IsArray(V) Then
'   If Sz(V) = 0 Then Exit Property
'End If
'IsEmp = False
'End Property
'
'Property Get IsFunTy(A$) As Boolean
'Select Case A
'Case "Property", "Sub", "Function": IsFunTy = True
'End Select
'End Property
'
'Property Get IsLetter(A) As Boolean
'Dim C1$: C1 = UCase(A)
'IsLetter = ("A" <= C1 And C1 <= "Z")
'End Property
'
'Property Get IsNm(A) As Boolean
'If Not IsLetter(FstChr(A)) Then Exit Property
'Dim L%: L = Len(A)
'If L > 64 Then Exit Property
'Dim J%
'For J = 2 To L
'   If Not IsNmChr(Mid(A, J, 1)) Then Exit Property
'Next
'IsNm = True
'End Property
'
'Property Get IsNmChr(A$) As Boolean
'IsNmChr = True
'If IsLetter(A) Then Exit Property
'If A = "_" Then Exit Property
'If IsDigit(A) Then Exit Property
'IsNmChr = False
'End Property
'Property Get MdDic(A As CodeModule) As Dictionary
'Set MdDic = Src_Dic_Of_MthKey_MthLines(MdSrc(A), MdPjNm(A), MdNm(A))
'End Property
'Sub ZZ_MdCmp()
'MdCmp Md("QTool.G_Tool"), Md("QSqTp.G_Tool")
'End Sub
'Sub MdCmp(A As CodeModule, B As CodeModule)
'Dim A1 As Dictionary, B1 As Dictionary
'    Set A1 = MdDic(A)
'    Set B1 = MdDic(B)
'Dim C As DCRslt
'    C = DicCmp(A1, B1, MdDNm(A), MdDNm(B))
'AyBrw DCRslt_Ly(C)
'End Sub
'Property Get IsNothing(A) As Boolean
'IsNothing = TypeName(A) = "Nothing"
'End Property
'Property Get IsMdNm(A) As Boolean
'Select Case Left(A, 2)
'Case "M_", "S_", "F_", "G_"
'    IsMdNm = True
'End Select
'End Property
'Property Get IsPfx(A, Pfx$) As Boolean
'IsPfx = Left(A, Len(Pfx)) = Pfx
'End Property
'
'Property Get IsPrim(A) As Boolean
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
'   IsPrim = True
'End Select
'End Property
'
'Property Get IsPun(C) As Boolean
'If IsLetter(C) Then Exit Property
'If IsDigit(C) Then Exit Property
'If C = "_" Then Exit Property
'IsPun = True
'End Property
'
'Property Get IsStr(A) As Boolean
'IsStr = VarType(A) = vbString
'End Property
'
'Property Get IsSy(A) As Boolean
'IsSy = VarType(A) = vbArray + vbString
'End Property
'
'Property Get ItrAy(A, OIntoAy)
'Dim O: O = OIntoAy: Erase O
'Dim I
'For Each I In A
'    Push O, I
'Next
'ItrAy = O
'End Property
'
'Property Get ItrNy(Itr) As String()
'Dim I, O$()
'For Each I In Itr
'    Push O, CallByName(I, "Name", VbGet)
'Next
'ItrNy = O
'End Property
'
'Property Get JnCrLf$(Ay)
'JnCrLf = Join(Ay, vbCrLf)
'End Property
'
'Property Get LasChr$(A)
'LasChr = Right(A, 1)
'End Property
'
'Property Get LinLCCOpt(L$, MthNm$, Lno%) As LCCOpt
'Dim A$
'Dim M$
'Dim N$
'A = LinRmvMdy(L)
'M = LinShiftMthTy(A)
'If M = "" Then Exit Property
'N = LinNm(A)
'If N <> MthNm Then Exit Property
'Dim C1%, C2%
'C1 = InStr(L, MthNm)
'C2 = C1 + Len(MthNm)
'With LinLCCOpt
'    .Som = True
'    With .LCC
'        .Lno = Lno
'        .C1 = C1
'        .C2 = C2
'    End With
'End With
'End Property
'
'Property Get LinFunTy$(MthLin$)
'Dim A$: A = LinRmvMdy(MthLin)
'Dim B$: B = LinT1(A)
'Select Case B
'Case "Function", "Sub", "Property": LinFunTy = B: Exit Property
'End Select
'End Property
'
'Property Get LinMdy$(L$)
'Dim A$
'A = "Private": If HasPfx(L, A) Then LinMdy = A: Exit Property
'A = "Friend":  If HasPfx(L, A) Then LinMdy = A: Exit Property
'A = "Public":  If HasPfx(L, A) Then LinMdy = A: Exit Property
'End Property
'
'Property Get LinNm$(A)
'Dim J%
'If Not IsLetter(Left(A, 1)) Then Exit Property
'For J = 2 To Len(A)
'    If Not IsNmChr(Mid(A, J, 1)) Then
'        LinNm = Left(A, J - 1)
'        Exit Property
'    End If
'Next
'LinNm = A
'End Property
'
'Property Get LinRmvMdy$(L$)
'Dim A$
'A = "": If HasPfx(L, A) Then LinRmvMdy = RmvPfx(L, A): Exit Property
'A = "Friend ":  If HasPfx(L, A) Then LinRmvMdy = RmvPfx(L, A): Exit Property
'A = "Public ":  If HasPfx(L, A) Then LinRmvMdy = RmvPfx(L, A): Exit Property
'LinRmvMdy = L
'End Property
'
'Property Get LinShiftMthTy$(O$)
'Dim A$
'A = "Property Get": If IsPfx(O, A) Then LinShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Property Let": If IsPfx(O, A) Then LinShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Property Set": If IsPfx(O, A) Then LinShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Function":     If IsPfx(O, A) Then LinShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'A = "Sub":          If IsPfx(O, A) Then LinShiftMthTy = A: O = Mid(O, Len(A) + 2): Exit Property
'End Property
'
'Property Get LinT1$(L)
'Dim A$: A = LTrim(L)
'Dim P%: P = InStr(A, " ")
'If P = 0 Then LinT1 = A: Exit Property
'LinT1 = Left(A, P - 1)
'End Property
'
'Property Get LinesAy_Wdt%(A$())
'Dim O%, J&, M%
'For J = 0 To UB(A)
'   M = LinesWdt(A(J))
'   If M > O Then O = M
'Next
'LinesAy_Wdt = O
'End Property
'
'Property Get LinesSqV(Lines$) As Variant
'LinesSqV = AySqV(SplitLines(Lines))
'End Property
'
'Property Get LinesTrimEnd$(A$)
'LinesTrimEnd = Join(LyTrimEnd(SplitLines(A)), vbCrLf)
'End Property
'
'Property Get LinesUnderLin$(Lines)
'LinesUnderLin = StrDup("-", LinesWdt(Lines))
'End Property
'
'Property Get LinesWdt%(A)
'LinesWdt = AyWdt(SplitLines(A))
'End Property
'
'Property Get LyTrimEnd(Ly) As String()
'If Sz(Ly) = 0 Then Exit Property
'Dim L$
'Dim J&
'For J = UB(Ly) To 0 Step -1
'    L = Trim(Ly(J))
'    If Trim(Ly(J)) <> "" Then
'        Dim O$()
'        O = Ly
'        ReDim Preserve O(J)
'        LyTrimEnd = O
'        Exit Property
'    End If
'Next
'End Property
'
'Property Get Max(A, B)
'If A > B Then
'    Max = A
'Else
'    Max = B
'End If
'End Property
'
'Property Get Md(MdDNm) As CodeModule
'Dim A$: A = MdDNm
'Dim P As VBProject
'Dim MdNm$
'    Dim L%
'    L = InStr(A, ".")
'    If L = 0 Then
'        Set P = CurPj
'        MdNm = A
'    Else
'        Dim PjNm$
'        PjNm = Left(A, L - 1)
'        Set P = Pj(PjNm)
'        MdNm = Mid(A, L + 1)
'    End If
'Set Md = P.VBComponents(MdNm).CodeModule
'End Property
'
'Property Get MdAllMthLinAy(A As CodeModule) As String()
'MdAllMthLinAy = SrcAllMthLinAy(MdSrc(A))
'End Property
'
'Sub MdClr(A As CodeModule, Optional IsSilent As Boolean)
'With A
'    If .CountOfLines = 0 Then Exit Sub
'    If Not IsSilent Then Debug.Print FmtQQ("MdClr: Md(?) of lines(?) is cleared", MdNm(A), .CountOfLines)
'    .DeleteLines 1, .CountOfLines
'End With
'End Sub
'
'Sub MdCpy(A As CodeModule, ToPj As VBProject)
'Dim MdNm$
'Dim FmPj As VBProject
'    Set FmPj = MdPj(A)
'    MdNm = A.Parent.Name
'If PjHasCmp(ToPj, MdNm) Then
'    Debug.Print FmtQQ("MdCpy: Md(?) exists in TarPj(?).  Skip moving", MdNm, ToPj.Name)
'    Exit Sub
'End If
'Dim TmpFil$
'    TmpFil = TmpFfn(".txt")
'    Dim SrcCmp As VBComponent
'    Set SrcCmp = A.Parent
'    SrcCmp.Export TmpFil
'    If SrcCmp.Type = vbext_ct_ClassModule Then
'        FtRmvFst4Lines TmpFil
'    End If
'Dim TarCmp As VBComponent
'    Set TarCmp = ToPj.VBComponents.Add(A.Parent.Type)
'    TarCmp.CodeModule.AddFromFile TmpFil
'Kill TmpFil
'PjSav ToPj
'Debug.Print FmtQQ("MdCpy: Md(?) is moved from SrcPj(?) to TarPj(?).", MdNm, FmPj.Name, ToPj.Name)
'End Sub
'
'Sub MdDlt(A As CodeModule)
'Dim M$, P$, Pj As VBProject
'    M = MdNm(A)
'    Set Pj = MdPj(A)
'    P = Pj.Name
'A.Parent.Collection.Remove A.Parent
'PjSav Pj
'Debug.Print FmtQQ("MdDlt: Md(?) is deleted from Pj(?)", M, P)
'End Sub
'
'Sub MdExport(A As CodeModule)
'Dim F$: F = MdSrcFfn(A)
'A.Parent.Export F
'Debug.Print MdNm(A)
'End Sub
'
'Property Get MdMthFNy(A As CodeModule, Optional NmPatn$ = ".", Optional IsNoSrt As Boolean) As String()
'Dim P$, M$, Sfx$, S$(), N$()
'    P = MdPjNm(A)
'    M = MdNm(A)
'    Sfx = ":" & P & "." & M
'    S = MdSrc(A)
'    N = SrcMthNy(S, NmPatn, IsNoSrt)
'MdMthFNy = AyAddSfx(N, Sfx)
'End Property
'
'Property Get Md_FunNy_OfPfx_ZZ_(A As CodeModule) As String()
'Dim J%, O$(), L$, L1$, Is_ZFun As Boolean
'For J = 1 To A.CountOfLines
'    Is_ZFun = True
'    L = A.Lines(J, 1)
'    Select Case True
'    Case IsPfx(L, "Sub ZZ_")
'        Is_ZFun = True
'        L1 = RmvPfx(L, "Sub ")
'    Case IsPfx(L, "Sub ZZ_")
'        Is_ZFun = True
'        L1 = RmvPfx(L, "Sub ")
'    Case Else:
'        Is_ZFun = False
'    End Select
'
'    If Is_ZFun Then
'        Push O, LinNm(L1)
'    End If
'Next
'Md_FunNy_OfPfx_ZZ_ = O
'End Property
'
'Sub MdGen_TstSub(A As CodeModule)
'MdRmv_TstSub A
'Dim Lines$: Lines = Md_TstSub_BdyLines(A)
'MdRmv_EmptyLines_AtEnd A
'If Lines <> "" Then
'    A.InsertLines A.CountOfLines + 1, Lines
'End If
'End Sub
'
'Sub MdGo(A As CodeModule)
'Cls_Win
'With A.CodePane
'    .Show
'    .Window.WindowState = vbext_ws_Maximize
'End With
'SendKeys "%WV"
'End Sub
'
'Sub MdGoLCCOpt(Md As CodeModule, LCCOpt As LCCOpt)
'MdGo Md
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
'Property Get MdHasTstSub(A As CodeModule) As Boolean
'Dim I
'For Each I In MdLy(A)
'    If I = "Friend Sub ZZ__Tst()" Then MdHasTstSub = True: Exit Property
'    If I = "Sub ZZ__Tst()" Then MdHasTstSub = True: Exit Property
'Next
'End Property
'
'Property Get MdIsAllRemarked(Md As CodeModule) As Boolean
'Dim J%, L$
'For J = 1 To Md.CountOfLines
'    If Left(Md.Lines(J, 1), 1) <> "'" Then Exit Property
'Next
'MdIsAllRemarked = True
'End Property
'
'Property Get MdIsMthBdy_Remarked(A As CodeModule, BdyFmToLno As FmToLno) As Boolean
'Dim B As FmToLno: B = BdyFmToLno
'Dim J%, Fm%
'Fm = B.FmLno
'If Not IsPfx(A.Lines(Fm, 1), "Stop '") Then Exit Property
'For J = Fm + 1 To B.ToLno
'    If Left(A.Lines(J, 1), 1) <> "'" Then Exit Property
'Next
'MdIsMthBdy_Remarked = True
'End Property
'
'Property Get MdLines$(A As CodeModule)
'With A
'    If .CountOfLines = 0 Then Exit Property
'    MdLines = .Lines(1, .CountOfLines)
'End With
'End Property
'
'Property Get MdLy(A As CodeModule) As String()
'MdLy = Split(MdLines(A), vbCrLf)
'End Property
'
'Sub MdMov_ToPj(A As CodeModule, ToPj As VBProject)
'If MdNm(A) = "F_Tool" And CurPj.Name = "QTool" Then
'    Debug.Print "Md(QTool.F_Tool) cannot be moved"
'    Exit Sub
'End If
'MdCpy A, ToPj
'MdDlt A
'End Sub
'
'Property Get MdMthDrs(A As CodeModule) As Drs
'Set MdMthDrs = Drs(SplitSsl(""), MdMthDry(A))
'End Property
'
'Property Get MdMthDry(A As CodeModule) As Variant()
'Dim O()
'MdMthDry = O
'End Property
'
'Property Get MdMthFmLno(A As CodeModule, MthNm$)
'MdMthFmLno = SrcMthFmLno(MdSrc(A), MthNm)
'End Property
'
'Property Get MdMthKy(A As CodeModule, Optional IsSngLinFmt As Boolean) As String()
'Dim PjN$: PjN = MdPjNm(A)
'Dim MdN$: MdN = MdNm(A)
'MdMthKy = SrcMthKy(MdSrc(A), PjN, MdN, IsSngLinFmt)
'End Property
'
'Property Get MdMthNy(A As CodeModule, Optional MthNmPatn$ = ".", Optional IsNoMdNmPfx As Boolean, Optional Mdy0$) As String()
'Dim Ay$(): Ay = SrcMthNy(MdSrc(A), MthNmPatn, Mdy0:=Mdy0)
'If IsNoMdNmPfx Then
'    MdMthNy = Ay
'Else
'    MdMthNy = AyAddPfx(Ay, MdNm(A) & ".")
'End If
'End Property
'Property Get DftPj(PjNm0$)
'If PjNm0 = "" Then
'    Set DftPj = CurPj
'Else
'    Set DftPj = Pj(PjNm0)
'End If
'End Property
'Property Get DftMd(MdDNm0$)
'If MdDNm0 = "" Then
'    Set DftMd = CurMd
'Else
'    Set DftMd = Md(MdDNm0)
'End If
'End Property
'Property Get DftMdDNm$(MdDNm0$)
'If MdDNm0 = "" Then
'    DftMdDNm = CurMdNm
'Else
'    DftMdDNm = MdDNm0
'End If
'End Property
'Property Get MdMthNy_OfInproper(A As CodeModule) As String()
'Dim MdN$: MdN = MdNm(A)
'    Dim Pfx$
'    Pfx = Left(MdN, 2)
'    If Pfx <> "M_" And Pfx <> "S_" Then
'        Debug.Print FmtQQ("MdMthNy_OfInproper: Given Md should be begins with [M_] or [S_].  MdNm=[?]", MdNm(A))
'        Exit Property ' M_Xxxx for Module with all pub-fun begins with Xxxx
'    End If                                             ' S_Xxxx for Module with single function of name=Xxxx
'Dim P$: P = Mid(MdN, 3) ' MthPfx
'Dim Ny$(), O$(), I
'Ny = MdMthNy(A, Mdy0:="Public", IsNoMdNmPfx:=True)
'PushAyNoDup Ny, MdMthNy(A, "ZZ_", IsNoMdNmPfx:=True)
'Ny = AyMinus(Ny, Array("ZZ__Tst"))
'If Sz(Ny) = 0 Then Exit Property
'Pfx = MdNm(A) & "."
'For Each I In Ny
'    If I <> "ZZ__Tst" Then
'        If Not IsPfx(I, P) Then Push O, Pfx & I
'    End If
'Next
'MdMthNy_OfInproper = O
'End Property
'
'Property Get MdNm$(A As CodeModule)
'MdNm = A.Parent.Name
'End Property
'
'Property Get MdPj(A As CodeModule) As VBProject
'Set MdPj = A.Parent.Collection.Parent
'End Property
'
'Property Get MdPjNm$(A As CodeModule)
'MdPjNm = MdPj(A).Name
'End Property
'
'Property Get MdRmk(A As CodeModule) As Boolean
'Debug.Print "Rmk " & A.Parent.Name,
'If MdIsAllRemarked(A) Then
'    Debug.Print " No need"
'    Exit Property
'End If
'Debug.Print "<============= is remarked"
'Dim J%
'For J = 1 To A.CountOfLines
'    A.ReplaceLine J, "'" & A.Lines(J, 1)
'Next
'MdRmk = True
'End Property
'
'Sub MdRmv_EmptyLines_AtEnd(A As CodeModule)
'Dim J%
'While A.CountOfLines > 1
'    J = J + 1
'    If J > 10000 Then Stop
'    If Trim(A.Lines(A.CountOfLines, 1)) <> "" Then Exit Sub
'    A.DeleteLines A.CountOfLines, 1
'Wend
'End Sub
'
'Sub MdRmv_TstSub(A As CodeModule)
'Dim L&, N&
'L = Md_TstSub_Lno(A)
'If L = 0 Then Exit Sub
'Dim Fnd As Boolean, J%
'For J = L + 1 To A.CountOfLines
'    If IsPfx(A.Lines(J, 1), "End Sub") Then
'        N = J - L + 1
'        Fnd = True
'        Exit For
'    End If
'Next
'If Not Fnd Then Stop
'A.DeleteLines L, N
'End Sub
'
'Property Get MdSrc(A As CodeModule) As String()
'MdSrc = MdLy(A)
'End Property
'
'Property Get MdSrcExt$(A As CodeModule)
'Dim O$
'Select Case A.Parent.Type
'Case vbext_ct_ClassModule: O = ".cls"
'Case vbext_ct_Document: O = ".cls"
'Case vbext_ct_StdModule: O = ".bas"
'Case vbext_ct_MSForm: O = ".cls"
'Case Else: Err.Raise 1, , "MdSrcExt: Unexpected MdCmpTy.  Should be [Class or Module or Document]"
'End Select
'MdSrcExt = O
'End Property
'
'Property Get MdSrcFfn$(A As CodeModule)
'MdSrcFfn = PjSrcPth(MdPj(A)) & MdSrcFn(A)
'End Property
'
'Property Get MdSrcFn$(A As CodeModule)
'MdSrcFn = MdNm(A) & MdSrcExt(A)
'End Property
'
'Sub MdSrt(A As CodeModule)
'If MdNm(A) = "F_Tool" And MdPjNm(A) = "QTool" Then
'    Exit Sub
'End If
'Dim Nm$: Nm = MdNm(A)
'Debug.Print "Sorting: "; AlignL(Nm, 30); " ";
'Dim Ay(): Ay = Array("M_A")
''Skip some md
'    If AyHas(Ay, Nm) Then
'        Debug.Print "<<<< Skipped"
'        Exit Sub
'    End If
'Dim NewLines$: NewLines = MdSrtedLines(A)
'Dim Old$: Old = MdLines(A)
''Exit if same
'    If Old = NewLines Then
'        Debug.Print "<== Same"
'        Exit Sub
'    End If
'Debug.Print "<-- Sorted";
''Delete
'    Debug.Print FmtQQ("<--- Deleted (?) lines", A.CountOfLines);
'    MdClr A, IsSilent:=True
''Add sorted lines
'    A.AddFromString NewLines
'    MdRmv_EmptyLines_AtEnd A
'    Debug.Print "<----Sorted Lines added...."
'End Sub
'
'Property Get MdSrtRpt(A As CodeModule) As DCRslt
'Dim P$, M$
'M = MdNm(A)
'P = MdPjNm(A)
'MdSrtRpt = SrcSrtRpt(MdSrc(A), P, M)
'End Property
'
'Property Get MdSrtRptLy(A As CodeModule) As String()
'Dim P$: P = MdPjNm(A)
'Dim M$: M = MdNm(A)
'MdSrtRptLy = SrcSrtRptLy(MdSrc(A), P, M)
'End Property
'
'Property Get MdSrtedLines$(A As CodeModule)
'MdSrtedLines = SrcSrtedLines(MdSrc(A))
'End Property
'
'Property Get Md_TstSub_BdyLines$(A As CodeModule)
'Dim Ny$(): Ny = Md_FunNy_OfPfx_ZZ_(A)
'If Sz(Ny) = 0 Then Exit Property
'Ny = AySrt(Ny)
'Dim O$()
'Dim Pfx$
'If A.Parent.Type = vbext_ct_ClassModule Then
'    Pfx = "Friend "
'End If
'Push O, ""
'Push O, Pfx & "Sub ZZ__Tst()"
'PushAy O, Ny
'Push O, "End Sub"
'Md_TstSub_BdyLines = Join(O, vbCrLf)
'End Property
'
'Property Get Md_TstSub_Lno%(A As CodeModule)
'Dim J%
'For J = 1 To A.CountOfLines
'    If SrcLin_IsTstSub(A.Lines(J, 1)) Then Md_TstSub_Lno = J: Exit Property
'Next
'End Property
'
'Property Get MdUnRmk(A As CodeModule) As Boolean
'Debug.Print "UnRmk " & A.Parent.Name,
'If Not MdIsAllRemarked(A) Then
'    Debug.Print "No need"
'    Exit Property
'End If
'Debug.Print "<===== is unmarked"
'Dim J%, L$
'For J = 1 To A.CountOfLines
'    L = A.Lines(J, 1)
'    If Left(L, 1) <> "'" Then Stop
'    A.ReplaceLine J, Mid(L, 2)
'Next
'MdUnRmk = True
'End Property
'
'Property Get Mdy_IsSel(A$, MdyAy$()) As Boolean
'If Sz(MdyAy) = 0 Then Mdy_IsSel = True: Exit Property
'Dim Mdy
'For Each Mdy In MdyAy
'    If Mdy = "Public" Then
'        If A = "" Then Mdy_IsSel = True: Exit Property
'    End If
'    If A = Mdy Then Mdy_IsSel = True: Exit Property
'Next
'End Property
'
'Property Get Min(ParamArray A())
'Dim O, J&, Av()
'Av = A
'O = A(0)
'For J = 1 To UB(Av)
'    If A(J) < O Then O = A(J)
'Next
'Min = O
'End Property
'
'Sub MthLin_BrkAsg(A$, Optional OIsMthLin As Boolean, Optional OMdy$, Optional OMajTy$, Optional OMthNm$)
'OMdy = LinMdy(A)
'OMthNm = ""
'OMajTy = ""
'
'Dim L$
'    If OMdy = "" Then L = A Else L = RmvPfx(A, OMdy & " ")
'
''OMajTy
'    Dim B$
'    B = "Sub ":          If HasPfx(L, B) Then L = RmvPfx(L, B): OMajTy = "Sub"
'    B = "Function ":     If HasPfx(L, B) Then L = RmvPfx(L, B): OMajTy = "Fun"
'    B = "Property Get ": If HasPfx(L, B) Then L = RmvPfx(L, B): OMajTy = "Prp"
'    B = "Property Let ": If HasPfx(L, B) Then L = RmvPfx(L, B): OMajTy = "Prp"
'    B = "Property Set ": If HasPfx(L, B) Then L = RmvPfx(L, B): OMajTy = "Prp"
'    If OMajTy = "" Then
'        OIsMthLin = False
'        Exit Sub
'    End If
'OMthNm = LinNm(L)
'OIsMthLin = True
'End Sub
'
'Property Get MthLin_MthKey$(A$, Optional PjNm$, Optional MdNm$, Optional IsSngLinFmt As Boolean)
'Dim M$ 'Mdy
'Dim T$ 'MthTy {Sub Fun Prp}
'Dim N$ 'Name
'Dim P% 'Priority
'    M = LinMdy(A)
'    Dim L$
'    If M = "" Then L = A Else L = RmvPfx(A, M & " ")
'    Dim B$
'    B = "Sub ":          If HasPfx(L, B) Then L = RmvPfx(L, B): T = "Sub"
'    B = "Function ":     If HasPfx(L, B) Then L = RmvPfx(L, B): T = "Fun"
'    B = "Property Get ": If HasPfx(L, B) Then L = RmvPfx(L, B): T = "Prp"
'    B = "Property Let ": If HasPfx(L, B) Then L = RmvPfx(L, B): T = "Prp"
'    B = "Property Set ": If HasPfx(L, B) Then L = RmvPfx(L, B): T = "Prp"
'    If T = "" Then Stop
'    N = LinNm(L)
'If IsPfx(N, "Init") And T = "Get" And M = "Friend" Then
'    P = 1
'ElseIf T = "Prp" And (M = "" Or M = "Public") Then
'    P = 2
'ElseIf HasSubStr(N, "__") Then
'    P = 4
'ElseIf N = "ZZ__Tst" Then
'    P = 9
'ElseIf IsPfx(N, "ZZ_") Then
'    P = 8
'ElseIf M = "Private" Then
'    P = 5
'Else
'    P = 3
'End If
'Dim F$, O$
'If PjNm = "" And MdNm = "" Then
'    F = IIf(IsSngLinFmt, "?:?:?:?", "?:?|?:?")
'    O = FmtQQ(F, P, N, T, M)
'
'Else
'    F = IIf(IsSngLinFmt, "?:?:?:?:?:?", "?:?|?:?|?:?")
'    O = FmtQQ(F, PjNm, MdNm, P, N, T, M)
'End If
'MthLin_MthKey = O
'End Property
'
'Property Get MthLin_MthNm$(A$)
'Dim N$ 'Name
'    MthLin_BrkAsg A, _
'        OMthNm:=N
'MthLin_MthNm = N
'End Property
'
'Sub MthDNm_Mov_ToProperMd(A)
'MthMovToProperMd MthDNm_Mth(A)
'End Sub
'
'Property Get MthDNm_Mth(A) As Mth
'Dim Ay$(): Ay = Split(A, ".")
'Dim Nm$, M As CodeModule
'Select Case Sz(Ay)
'Case 1: Nm = Ay(0): Set M = CurMd
'Case 2: Nm = Ay(1): Set M = Md(A)
'Case 3: Nm = Ay(2): Set M = Md(Ay(0) & "." & Ay(1))
'End Select
'Set MthDNm_Mth = Mth(M, Nm)
'End Property
'
'Property Get MthBdyFmToLno(A As Mth) As FmToLno
'MthBdyFmToLno = SrcMthBdyFmToLno(MdSrc(A.Md), A.Nm)
'End Property
'
'Sub MthBrkAsg(A As Mth, OMdy$, OFunTy$)
'Dim L$: L = MthLin(A)
'OMdy = SrcLin_Mdy(L)
'OFunTy = SrcLin_FunTy(L)
'End Sub
'
'Sub MthCpy(A As Mth, ToMd As CodeModule, Optional IsSilent As Boolean)
'If MdHasMth(ToMd, A.Nm) Then
'    Debug.Print FmtQQ("MthCpy_ToMd: Fm-Mth(?) is Found in To-Md(?)", A.Nm, MdNm(ToMd))
'    Exit Sub
'End If
'If ObjPtr(A.Md) = ObjPtr(ToMd) Then
'    Debug.Print FmtQQ("MthCpy_ToMd: Fm-Mth-Md(?) cannot be To-Md(?)", MthMdNm(A), MdNm(ToMd))
'    Exit Sub
'End If
'MdAppLines ToMd, MthLines(A)
'If Not IsSilent Then
'    Debug.Print FmtQQ("MthCpy_ToMd: Mth(?) is copied ToMd(?)", MthDNm(A), MdDNm(ToMd))
'End If
'End Sub
'
'Sub MthGo(A As Mth)
'MdGoLCCOpt A.Md, MthLCCOpt(A)
'End Sub
'
'Property Get MthIsExist(A As Mth) As Boolean
'MthIsExist = MdMthFmLno(A.Md, A.Nm) > 0
'End Property
'
'Property Get MthIsPub(A As Mth) As Boolean
'Dim L$: L = MthLin(A)
'If L = "" Then Stop
'Dim Mdy$: Mdy = SrcLin_Mdy(L)
'If Mdy = "" Or Mdy = "Public" Then MthIsPub = True
'End Property
'
'Property Get MthLCCOpt(A As Mth) As LCCOpt
'Dim L%, C As LCCOpt
'Dim M As CodeModule
'Set M = A.Md
'For L = M.CountOfDeclarationLines + 1 To M.CountOfLines
'    C = LinLCCOpt(M.Lines(L, 1), A.Nm, L)
'    If C.Som Then
'        MthLCCOpt.Som = True
'        MthLCCOpt = C
'        Exit Property
'    End If
'Next
'Stop
'End Property
'
'Property Get MthLin$(A As Mth)
'MthLin = SrcMthLin(MdSrc(A.Md), A.Nm)
'End Property
'
'Property Get MthLines$(A As Mth)
'MthLines = Src_MthLines_ByNm(MdSrc(A.Md), A.Nm)
'End Property
'Property Get MdHasMth(A As CodeModule, MthNm$) As Boolean
'MdHasMth = MdMthFmLno(A, MthNm) > 0
'End Property
'
'Property Get MthMdNm$(A As Mth)
'MthMdNm = MdNm(A.Md)
'End Property
'
'Sub MthMov(A As Mth, ToMd As CodeModule)
'MthCpy A, ToMd
'MthRmv A
'End Sub
'
'Sub MthMovToProperMd(A As Mth)
'MthMov A, MthProperMd(A)
'End Sub
'Property Get DftMth(MthDNm0$) As Mth
'If MthDNm0 = "" Then
'    Set DftMth = CurMth
'    Exit Property
'End If
'Set DftMth = MthDNm_Mth(MthDNm0)
'End Property
'
'
'Sub MthRmk(A As Mth)
'Dim P As FmToLno
'    P = MthBdyFmToLno(A)
'Dim M As CodeModule: Set M = A.Md
'If MdIsMthBdy_Remarked(M, P) Then Exit Sub
'Dim J%, L$
'For J = P.FmLno To P.ToLno
'    L = M.Lines(J, 1)
'    M.ReplaceLine J, "'" & L
'Next
'M.InsertLines P.FmLno, "Stop" & " '"
'End Sub
'
'Sub MthRmv(A As Mth, Optional IsSilent As Boolean)
'Dim J%, FmLno%, ToLno%, Cnt%, S$(), L%()
'S = MdSrc(A.Md)
'L = SrcMthFmLnoAy(S, A.Nm)
'For J = UB(L) To 0 Step -1
'    FmLno = L(J)
'    ToLno = SrcMthToLno(S, FmLno)
'    Cnt = ToLno - FmLno + 1
'    A.Md.DeleteLines FmLno, Cnt
'Next
'If Not IsSilent Then
'    Debug.Print FmtQQ("MthRmv: Mth(?) of LinCnt(?) is deleted", MthDNm(A), Cnt)
'End If
'End Sub
'
'Sub MthUnRmk(A As Mth)
'Dim P As FmToLno
'    P = MthBdyFmToLno(A)
'Dim M As CodeModule: Set M = A.Md
'If Not MdIsMthBdy_Remarked(M, P) Then Exit Sub
'Dim J%, L$
'For J = P.FmLno + 1 To P.ToLno
'    L = M.Lines(J, 1)
'    If Left(L, 1) <> "'" Then Stop
'    M.ReplaceLine J, Mid(L, 2)
'Next
'If Not IsPfx(M.Lines(P.FmLno, 1), "Stop '") Then Stop
'M.DeleteLines P.FmLno, 1
'End Sub
'
'Property Get NewWb() As Workbook
'Set NewWb = Xls.Workbooks.Add
'End Property
'
'Property Get NewWs() As Worksheet
'Set NewWs = NewWb.Sheets(1)
'End Property
'
'Sub OyDo(Oy, DoFunNm$)
'Dim O
'For Each O In Oy
'    Excel.Run DoFunNm, O ' DoFunNm cannot be like a Excel.Address (eg, A1, XX1)
'Next
'End Sub
'
'Property Get OyNy(Oy) As String()
'Dim O$(): If Sz(Oy) = 0 Then Exit Property
'Dim I
'For Each I In Oy
'    Push O, CallByName(I, "Name", VbGet)
'Next
'OyNy = O
'End Property
'
'Property Get Pj(PjNm$) As VBProject
'Set Pj = CurVbe.VBProjects(PjNm)
'End Property
'
'Property Get PjMbrDotNm_Either(A) As Either
''Return ~.Left as PjMbrDotNm
''Or     ~.Right as PjNy() for those Pj holding giving Md
'Dim P$, M$
'Brk2_Asg A, ".", P, M
'If P <> "" Then
'    PjMbrDotNm_Either = EitherL(A)
'    Exit Property
'End If
'Dim Ny$()
'Ny = CurVbe_MdPjNy(M)
'If Sz(Ny) = 1 Then
'    PjMbrDotNm_Either = EitherL(Ny(0) & "." & M)
'    Exit Property
'End If
'PjMbrDotNm_Either = EitherR(Ny)
'End Property
'
'Sub PjAddRf(A As VBProject, RfNm$)
'Dim RfFfn$: RfFfn = RfNm_RfFfn(RfNm)
'If RfFfn = "" Then Stop
'Dim F$: F = PjFfn(A)
'If F = "" Then Exit Sub
'If F = RfFfn Then Exit Sub
'If PjHasRfNm(A, RfNm) Then Exit Sub
'A.References.AddFromFile RfFfn
'PjSav A
'End Sub
'
'Sub PjAddCls(A As VBProject, Nm$)
'PjAddMbr A, Nm, vbext_ct_ClassModule
'End Sub
'
'Sub PjAddMbr(A As VBProject, Nm$, Ty As vbext_ComponentType, Optional IsGoMbr As Boolean)
'If PjHasCmp(A, Nm) Then
'    MsgBox FmtQQ("Cmp(?) exist in CurPj(?)", Nm, CurPjNm), , "M_A.ZAddMbr"
'    Exit Sub
'End If
'Dim Cmp As VBComponent
'Set Cmp = A.VBComponents.Add(Ty)
'Cmp.Name = Nm
'Cmp.CodeModule.InsertLines 1, "Option Explicit"
'If IsGoMbr Then Shw_Mbr Nm
'End Sub
'
'Property Get PjClsNy_With_TstSub(A As VBProject) As String()
'Dim I As VBComponent
'Dim O$()
'For Each I In A.VBComponents
'    If I.Type = vbext_ct_ClassModule Then
'        If MdHasTstSub(I.CodeModule) Then
'            Push O, I.Name
'        End If
'    End If
'Next
'PjClsNy_With_TstSub = O
'End Property
'
'Property Get PjCmp(A As VBProject, Nm) As VBComponent
'Set PjCmp = A.VBComponents(CStr(Nm))
'End Property
'
'Sub PjCompile(A As VBProject)
'PjGo A
'SendKeys "%D{Enter}"
'End Sub
'
'Sub PjCrt_Fxa(A As VBProject, FxaNm$)
'Dim F$
'F = FxaNm_Fxa(FxaNm)
'End Sub
'
'Property Get PjDupMthFNy(A As VBProject, Optional IsNoSrt As Boolean, Optional IsSamMthBdyOnly As Boolean) As String()
'Dim N$(): N = PjMthFNy(A, IsNoSrt:=IsNoSrt)
'Dim N1$(): N1 = MthFNy_DupMthFNy(N)
'If IsSamMthBdyOnly Then
'    N1 = DupMthFNy_SamMthBdyMthFNy(N1, A)
'End If
'PjDupMthFNy = N1
'End Property
'
'Sub PjEns_Cls(A As VBProject, ClsNm$)
'PjEns_Cmp A, ClsNm, vbext_ct_ClassModule
'End Sub
'Sub MdRpl_Cxt(A As CodeModule, Cxt$)
'Dim N%: N = A.CountOfLines
'MdClr A, IsSilent:=True
'A.AddFromString Cxt
'Debug.Print FmtQQ("MdRpl_Cxt: Md(?) of Ty(?) of Old-LinCxt(?) is replaced by New-Len(?) New-LinCnt(?).<-----------------", _
'    MdDNm(A), MdTyNm(A), N, Len(Cxt), LinesLinCnt(Cxt))
'End Sub
'Property Get MdTyNm$(A As CodeModule)
'MdTyNm = CmpTy_Nm(MdCmpTy(A))
'End Property
'Property Get StrSubStrCnt&(A$, SubStr$)
'Dim P&, O%, L%
'L = Len(SubStr)
'P = 1
'While P > 0
'    P = InStr(P, A, SubStr)
'    If P > 0 Then O = O + 1: P = P + L
'Wend
'StrSubStrCnt = O
'End Property
'Property Get LinesLinCnt%(A$)
'LinesLinCnt = StrSubStrCnt(A, vbCrLf) + 1
'End Property
'Property Get MdCmpTy(A As CodeModule) As vbext_ComponentType
'MdCmpTy = A.Parent.Type
'End Property
'
'Sub PjEns_Cmp(A As VBProject, Nm$, Ty As vbext_ComponentType)
'If PjHasCmp(A, Nm) Then Exit Sub
'Dim Cmp As VBComponent
'Set Cmp = A.VBComponents.Add(Ty)
'Cmp.Name = Nm
'Cmp.CodeModule.InsertLines 1, "Option Explicit"
'Debug.Print FmtQQ("PjEns_Cmp: Md(?) of Ty(?) is added in Pj(?) <===================================", Nm, CmpTy_Nm(Ty), A.Name)
'End Sub
'
'Sub PjEns_Md(A As VBProject, MdNm$)
'PjEns_Cmp A, MdNm, vbext_ct_StdModule
'End Sub
'
'Sub PjExport(A As VBProject)
'Dim P$: P = PjSrcPth(A)
'If P = "" Then
'    Debug.Print FmtQQ("PjExport: Pj(?) does not have FileName", A.Name)
'    Exit Sub
'End If
'PthClrFil P 'Clr SrcPth ---
'FfnCpyToPth A.Filename, P, OvrWrt:=True
'Dim I, Ay() As CodeModule
'Ay = PjMbrAy(A)
'If Sz(Ay) = 0 Then Exit Sub
'For Each I In Ay
'    MdExport CvMd(I)  'Exp each md --
'Next
'AyWrt PjRfLy(A), PjRfCfgFfn(A) 'Exp rf -----
'End Sub
'
'Property Get PjMthFNy(A As VBProject, Optional IsNoSrt As Boolean) As String()
'Dim Ay() As CodeModule
'    Ay = PjMdAy(A)
'If Sz(Ay) = 0 Then Exit Property
'Dim O$(), I
'For Each I In Ay
'    PushAy O, MdMthFNy(CvMd(I), IsNoSrt:=True)
'Next
'If IsNoSrt Then
'    PjMthFNy = O
'Else
'    PjMthFNy = AySrt(O)
'End If
'End Property
'
'Property Get PjFfn$(A As VBProject)
'On Error Resume Next
'PjFfn = A.Filename
'End Property
'
'Property Get PjFstMd(A As VBProject) As CodeModule
'Dim Cmp As VBComponent, O$()
'For Each Cmp In A.VBComponents
'    If Cmp.Type = vbext_ct_StdModule Then
'        Set PjFstMd = Cmp.CodeModule
'        Exit Property
'    End If
'Next
'For Each Cmp In A.VBComponents
'    If Cmp.Type = vbext_ct_ClassModule Then
'        Set PjFstMd = Cmp.CodeModule
'        Exit Property
'    End If
'Next
'End Property
'
'Property Get PjFunBdyDic(A As VBProject) As Dictionary
'Stop '
'End Property
'
'Property Get PjFunNy(A As VBProject, Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".") As String()
'Dim Ay() As CodeModule: Ay = PjMbrAy(A, MbrNmPatn)
'If Sz(Ay) = 0 Then Exit Property
'Dim I, O$()
'For Each I In Ay
'    PushAy O, MdMthNy(CvMd(I), MthNmPatn)
'Next
'O = AyAddPfx(O, A.Name & ".")
'PjFunNy = O
'End Property
'
'Sub Pj_Gen_TstClass(A As VBProject)
'If PjHasCmp(A, "Tst") Then
'    CmpRmv PjCmp(A, "Tst")
'End If
'PjAddCls A, "Tst"
'PjMd(A, "Tst").AddFromString Pj_TstClass_Bdy(A)
'End Sub
'
'Sub Pj_Gen_TstSub(A As VBProject)
'Dim Ny$(): Ny = PjMd_and_Cls_Ny(A)
'Dim N, M As CodeModule
'For Each N In Ny
'    Set M = A.VBComponents(N).CodeModule
'    MdGen_TstSub M
'Next
'End Sub
'
'Sub PjGo(A As VBProject)
'Cls_Win
'Dim Md As CodeModule
'Set Md = PjFstMd(A)
'If IsNothing(Md) Then Exit Sub
'Md.CodePane.Show
'SendKeys "%WV" ' Window SplitVertical
'End Sub
'
'Property Get PjHasCmp(A As VBProject, Nm$) As Boolean
'Dim Cmp As VBComponent
'For Each Cmp In A.VBComponents
'    If Cmp.Name = Nm Then PjHasCmp = True: Exit Property
'Next
'End Property
'
'Property Get PjHasRfNm(A As VBProject, RfNm$) As Boolean
'Dim I, R As Reference
'For Each I In A.References
'    Set R = I
'    If R.Name = RfNm Then PjHasRfNm = True: Exit Property
'Next
'End Property
'
'Property Get PjMbrAy(A As VBProject, Optional MbrNmPatn$ = ".") As CodeModule()
'PjMbrAy = PjMbrAy__X(A, MbrNmPatn, CmpTyAy_Of_Cls_and_Md)
'End Property
'
'Property Get PjMbrNy(A As VBProject, Optional MbrNmPatn$ = ".") As String()
'PjMbrNy = OyNy(PjMbrAy(A, MbrNmPatn))
'End Property
'
'Property Get PjMd(A As VBProject, Nm) As CodeModule
'Set PjMd = PjCmp(A, Nm).CodeModule
'End Property
'
'Property Get PjMdAy(A As VBProject, Optional MdNmPatn$ = ".") As CodeModule()
'PjMdAy = PjMbrAy__X(A, MdNmPatn, CmpTyAy_Of_Md)
'End Property
'
'Property Get PjMdNy_With_TstSub(A As VBProject) As String()
'Dim I As VBComponent
'Dim O$()
'For Each I In A.VBComponents
'    If I.Type = vbext_ct_StdModule Then
'        If MdHasTstSub(I.CodeModule) Then
'            Push O, I.Name
'        End If
'    End If
'Next
'PjMdNy_With_TstSub = O
'End Property
'
'Property Get PjMdSrtRpt(A As VBProject) As MdSrtRpt
''SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
'Dim Ay() As CodeModule: Ay = PjMbrAy(A)
'Dim Ny$(): Ny = OyNy(Ay)
'Dim LyAy()
'Dim IsSam() As Boolean
'    Dim J%, R As DCRslt
'    For J = 0 To UB(Ay)
'        R = MdSrtRpt(Ay(J))
'        Push LyAy, DCRslt_Ly(R)
'        Push IsSam, DCRslt_IsSam(R)
'    Next
'With PjMdSrtRpt
'    Set .RptDic = AyPair_Dic(Ny, LyAy)
'    .MdNy = PjMdSrtRpt_1(Ny, IsSam)
'End With
'End Property
'
'Property Get PjMdSrtRpt_1(MdNy$(), IsSam() As Boolean) As String()
'Dim O$(), J%
'For J = 0 To UB(MdNy)
'    Push O, AlignL(MdNy(J), 30) & " " & IsSam(J)
'Next
'PjMdSrtRpt_1 = O
'End Property
'
'Property Get PjMd_and_Cls_Ny(A As VBProject) As String()
'Dim O$(), Cmp As VBComponent
'For Each Cmp In A.VBComponents
'    If Cmp.Type = vbext_ct_StdModule Or Cmp.Type = vbext_ct_ClassModule Then
'        Push O, Cmp.Name
'    End If
'Next
'PjMd_and_Cls_Ny = O
'End Property
'
'Property Get PjMthKy(A As VBProject, Optional IsSngLinFmt As Boolean) As String()
'Dim O$(), I
'Dim Ay() As CodeModule
'Ay = PjMbrAy(A)
'If Sz(Ay) = 0 Then Exit Property
'For Each I In Ay
'    PushAy O, MdMthKy(CvMd(I), IsSngLinFmt)
'Next
'PjMthKy = O
'End Property
'
'Property Get PjMthNy(A As VBProject, Optional MthNmPatn$ = ".", Optional MbrNmPatn$ = ".", Optional Mdy0$) As String()
'Dim Ay() As CodeModule: Ay = PjMbrAy(A, MbrNmPatn)
'If Sz(Ay) = 0 Then Exit Property
'Dim I, O$()
'For Each I In Ay
'    PushAy O, MdMthNy(CvMd(I), MthNmPatn, Mdy0:=Mdy0)
'Next
'O = AyAddPfx(O, A.Name & ".")
'PjMthNy = O
'End Property
'
'Property Get PjMthNy_OfInproper(A As VBProject) As String()
'Dim I, O$()
'Dim Ay() As CodeModule: Ay = PjMdAy(A)
'If Sz(Ay) = 0 Then Exit Property
'For Each I In Ay
'    Dim Pfx$: Pfx = Left(MdNm(CvMd(I)), 2)
'    If Pfx = "M_" Or Pfx = "S_" Then
'        PushAy O, MdMthNy_OfInproper(CvMd(I))
'    End If
'Next
'PjMthNy_OfInproper = AyAddPfx(O, A.Name & ".")
'End Property
'
'Property Get Pj_Dic_Of_MthKey_MthLines(A As VBProject) As Dictionary
'Dim I
'Dim O As New Dictionary
'For Each I In PjMbrAy(A)
'    Set O = DicAdd(O, Md_Dic_Of_MthKey_MthLines(CvMd(I)))
'Next
'Set Pj_Dic_Of_MthKey_MthLines = O
'End Property
'
'Property Get PjPth$(A As VBProject)
'PjPth = FfnPth(A.Filename)
'End Property
'
'Property Get PjRfAy(A As VBProject) As Reference()
'PjRfAy = ItrAy(A.References, EmpRfAy)
'End Property
'
'Property Get PjRfCfgFfn(A As VBProject)
'PjRfCfgFfn = PjSrcPth(A) & "PjRf.Cfg"
'End Property
'
'Property Get PjRfLy(A As VBProject) As String()
'Dim RfAy() As Reference
'    RfAy = PjRfAy(A)
'Dim O$()
'Dim Ny$(): Ny = OyNy(RfAy)
'Ny = AyAlignL(Ny)
'Dim J%
'For J = 0 To UB(Ny)
'    Push O, Ny(J) & " " & RfFfn(RfAy(J))
'Next
'PjRfLy = O
'End Property
'
'Sub PjSav(A As VBProject)
'PjGo A
'SendKeys "^S"
'End Sub
'
'Property Get PjSrcPth(A As VBProject)
'Dim Ffn$: Ffn = PjFfn(A)
'If Ffn = "" Then Exit Property
'Dim Fn$: Fn = FfnFn(Ffn)
'Dim P$: P = FfnPth(A.Filename)
'If P = "" Then Exit Property
'Dim O$:
'O = P & "Src\": PthEns O
'O = O & Fn & "\":                  PthEns O
'PjSrcPth = O
'End Property
'
'Sub PjSrcPthBrw(A As VBProject)
'PthBrw PjSrcPth(A)
'End Sub
'
'Sub PjSrt(A As VBProject)
'If A.Name = "QTool" Then Exit Sub
'Dim I
'Dim Ny$(): Ny = AySrt(PjMd_and_Cls_Ny(A))
'If Sz(Ny) = 0 Then Exit Sub
'For Each I In Ny
'    MdSrt PjMd(A, I)
'Next
'End Sub
'
'Property Get PjSrtRptLy(A As VBProject) As String()
'Dim Ay() As CodeModule: Ay = PjMbrAy(A)
'Dim O$(), I, M As CodeModule
'For Each I In Ay
'    Set M = I
'    PushAy O, MdSrtRptLy(M)
'Next
'PjSrtRptLy = O
'End Property
'
'Function PjSrtRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
'Dim A1 As MdSrtRpt
'A1 = PjMdSrtRpt(A)
'Dim O As Workbook: Set O = DicWb(A1.RptDic)
'Dim Ws As Worksheet
'Set Ws = WbAddWs(O, "Md Idx")
''Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
''LoCol_LnkWs Lo, "Md"
''If Vis Then WbVis O
''Set PjSrtRptWb = O
'Stop '
'End Function
'
'Property Get Pj_TstClass_Bdy$(A As VBProject)
'Dim N1$() ' All Class Ny with 'Friend Sub ZZ__Tst' method
'Dim N2$()
'Dim A1$, A2$
'Const Q1$ = "Sub ?()|Dim A As New ?: A.ZZ__Tst|End Sub"
'Const Q2$ = "Sub ?()|#.?.ZZ__Tst|End Sub"
'N1 = PjClsNy_With_TstSub(A)
'A1 = SeedExpand(Q1, N1)
'N2 = PjMdNy_With_TstSub(A)
'A2 = Replace(SeedExpand(Q2, N2), "#", A.Name)
'Pj_TstClass_Bdy = A1 & vbCrLf & A2
'End Property
'
''Function FfnRplExt$(Ffn, NewExt)
''FfnRplExt = FfnRmvExt(Ffn) & NewExt
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
'Sub PthBrw(P)
'Shell "Explorer """ & P & """", vbMaximizedFocus
'End Sub
'
'Sub PthClrFil(A)
'Dim F
'For Each F In PthFfnColl(A)
'   FfnDlt F
'Next
'End Sub
'
'Sub PthEns(P$)
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
'Property Get PthFfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
'PthFfnAy = AyAddPfx(PthFnAy(A, Spec, Atr), A)
'End Property
'
'Property Get PthFfnColl(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As Collection
'Set PthFfnColl = CollAddPfx(PthFnColl(A, Spec, Atr), A)
'End Property
'
'Property Get PthFnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
'If Not PthIsExist(A) Then
'    Debug.Print FmtQQ("PthFnAy: Given Path(?) does not exit", A)
'    Exit Property
'End If
'Dim O$()
'Dim M$
'M = Dir(A & Spec)
'If Atr = 0 Then
'    While M <> ""
'       Push O, M
'       M = Dir
'    Wend
'    PthFnAy = O
'End If
'Ass PthHasPthSfx(A)
'While M <> ""
'    If GetAttr(A & M) And Atr Then
'        Push O, M
'    End If
'    M = Dir
'Wend
'PthFnAy = O
'End Property
'
'Property Get PthFnColl(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As Collection
'Set PthFnColl = AyColl(PthFnAy(A, Spec, Atr))
'End Property
'
'Property Get PthHasPthSfx(A) As Boolean
'PthHasPthSfx = LasChr(A) = "\"
'End Property
'
'Property Get PthIsExist(A) As Boolean
'Ass PthHasPthSfx(A)
'PthIsExist = Fso.FolderExists(A)
'End Property
'
'Sub Push(O, M)
'Dim N&
'    N = Sz(O)
'ReDim Preserve O(N)
'If IsObject(M) Then
'    Set O(N) = M
'Else
'    O(N) = M
'End If
'End Sub
'
'Sub PushAy(OAy, Ay)
'If Sz(Ay) = 0 Then Exit Sub
'Dim I
'For Each I In Ay
'    Push OAy, I
'Next
'End Sub
'Sub PushAyNoDup(OAy, Ay)
'If Sz(Ay) = 0 Then Exit Sub
'Dim I
'For Each I In Ay
'    PushNoDup OAy, I
'Next
'End Sub
'
'Sub PushNoDup(O, M)
'If Not AyHas(O, M) Then Push O, M
'End Sub
'
'Sub PushNonEmp(O, M)
'If IsEmp(M) Then Exit Sub
'Push O, M
'End Sub
'
'Sub PushObj(O, M)
'If Not IsObject(M) Then Stop
'Dim N&
'    N = Sz(O)
'ReDim Preserve O(N)
'Set O(N) = M
'End Sub
'
'Property Get Re(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
'Dim O As New RegExp
'With O
'   .Pattern = Patn
'   .MultiLine = MultiLine
'   .IgnoreCase = IgnoreCase
'   .Global = IsGlobal
'End With
'Set Re = O
'End Property
'
'Property Get RfNm_RfFfn$(RfNm$)
'Dim Ay() As VBProject: Ay = CurVbe_PjAy
'Dim M As VBProject, I
'For Each I In Ay
'    Set M = I
'    If M.Name = RfNm Then RfNm_RfFfn = M.Filename: Exit Property
'Next
'End Property
'
'Property Get RfFfn$(A As Reference)
'On Error Resume Next
'RfFfn = A.FullPath
'End Property
'
'Property Get RgRC(A As Range, R, C) As Range
'Set RgRC = A.Cells(R, C)
'End Property
'
'Property Get RgRCRC(A As Range, R1, C1, R2, C2) As Range
'Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
'End Property
'
'Property Get RgWs(A As Range)
'Set RgWs = A.Parent
'End Property
'
'Property Get RmvLasChr$(A)
'RmvLasChr = Left(A, Len(A) - 1)
'End Property
'
'Function RmvPfx$(A, Pfx$)
'If IsPfx(A, Pfx) Then
'    RmvPfx = Mid(A, Len(Pfx) + 1)
'Else
'    RmvPfx = A
'End If
'End Function
'
'Property Get RplDblSpc$(A)
'Dim O$: O = Trim(A)
'Dim J&
'While HasSubStr(O, "  ")
'    J = J + 1: If J > 10000 Then Stop
'    O = Replace(O, "  ", " ")
'Wend
'RplDblSpc = O
'End Property
'
'Property Get RplPun$(A)
'Dim O$(), J&, L&, C$
'L = Len(A)
'If L = 0 Then Exit Property
'ReDim O(L - 1)
'For J = 1 To L
'    C = Mid(A, J, 1)
'    If IsPun(C) Then
'        O(J - 1) = " "
'    Else
'        O(J - 1) = C
'    End If
'Next
'RplPun = Join(O, "")
'End Property
'
'Property Get RplVBar$(A)
'RplVBar = Replace(A, "|", vbCrLf)
'End Property
'
'Property Get S1S2Ay_Add(A() As S1S2, B() As S1S2) As S1S2()
'Dim O() As S1S2
'Dim J&
'O = A
'For J = 0 To UB(B)
'    PushObj O, B(J)
'Next
'S1S2Ay_Add = O
'End Property
'
'Sub S1S2Ay_Brw(A() As S1S2)
'AyBrw S1S2Ay_FmtLy(A)
'End Sub
'
'Property Get S1S2Ay_Dic(A() As S1S2) As Dictionary
'Dim J&, O As New Dictionary
'For J = 0 To UB(A)
'    O.Add A(J).S1, A(J).S2
'Next
'Set S1S2Ay_Dic = O
'End Property
'Property Get S1S2Ay_FmtLy(A() As S1S2) As String()
'Dim W1%: W1 = S1S2Ay_S1LinesWdt(A)
'Dim W2%: W2 = S1S2Ay_S2LinesWdt(A)
'Dim W%(1)
'W(0) = W1
'W(1) = W2
'Dim H$: H = WdtAy_HdrLin(W)
'S1S2Ay_FmtLy = S1S2Ay_LinesLinesLy(A, H, W1, W2)
'End Property
'
'Property Get S1S2Ay_LinesLinesLy(A() As S1S2, H$, W1%, W2%) As String()
'Dim O$(), I&
'Push O, H
'For I = 0 To UB(A)
'   PushAy O, S1S2_Ly(A(I), W1, W2)
'   Push O, H
'Next
'S1S2Ay_LinesLinesLy = O
'End Property
'
'Property Get S1S2Ay_S1LinesWdt%(A() As S1S2)
'S1S2Ay_S1LinesWdt = LinesAy_Wdt(S1S2Ay_Sy1(A))
'End Property
'
'Property Get S1S2Ay_S2LinesWdt%(A() As S1S2)
'S1S2Ay_S2LinesWdt = LinesAy_Wdt(S1S2Ay_Sy2(A))
'End Property
'
'Property Get S1S2Ay_Sy1(A() As S1S2) As String()
'Dim O$(), J&
'For J = 0 To UB(A)
'   Push O, A(J).S1
'Next
'S1S2Ay_Sy1 = O
'End Property
'
'Property Get S1S2Ay_Sy2(A() As S1S2) As String()
'Dim O$(), J&
'For J = 0 To UB(A)
'   Push O, A(J).S2
'Next
'S1S2Ay_Sy2 = O
'End Property
'
'Property Get S1S2_Ly(A As S1S2, W1%, W2%) As String()
'Dim S1$(), S2$()
'S1 = SplitLines(A.S1)
'S2 = SplitLines(A.S2)
'Dim M%, J%, O$(), Lin$, A1$, A2$, U1%, U2%
'    U1 = UB(S1)
'    U2 = UB(S2)
'    M = Max(U1, U2)
'Dim Spc1$, Spc2$
'    Spc1 = Space(W1)
'    Spc2 = Space(W2)
'For J = 0 To M
'   If J > U1 Then
'       A1 = Spc1
'   Else
'       A1 = StrAlignL(S1(J), W1)
'   End If
'   If J > U2 Then
'       A2 = Spc2
'   Else
'       A2 = StrAlignL(S2(J), W2)
'   End If
'   Lin = "| " + A1 + " | " + A2 + " |"
'   Push O, Lin
'Next
'S1S2_Ly = O
'End Property
'
'Property Get SeedExpand$(QVbl$, Ny$())
'Dim O$()
'Dim Sy$(): Sy = SplitVBar(QVbl)
'Dim J%, I
'For J = 0 To UB(Ny)
'    For Each I In Sy
'       Push O, Replace(I, "?", Ny(J))
'    Next
'Next
'SeedExpand = JnCrLf(O)
'End Property
'
'Property Get SplitLines(A) As String()
'Dim B$: B = Replace(A, vbCrLf, vbLf)
'SplitLines = Split(B, vbLf)
'End Property
'
'Property Get SplitSsl(A) As String()
'SplitSsl = Split(RplDblSpc(Trim(A)), " ")
'End Property
'
'Property Get SplitVBar(Vbl$) As String()
'SplitVBar = Split(Vbl, "|")
'End Property
'
'Property Get SrcLin_FunTy$(A)
'Dim A1$, A2$
'A1 = SrcLin_RmvMdy(A)
'A2 = LinT1(A1)
'If IsFunTy(A2) Then SrcLin_FunTy = A2
'End Property
'
'Property Get SrcLin_IsCd(A) As Boolean
'Dim L$: L = Trim(A)
'If A = "" Then Exit Property
'If Left(A, 1) = "'" Then Exit Property
'SrcLin_IsCd = True
'End Property
'
'Property Get SrcLin_IsMth(A) As Boolean
'SrcLin_IsMth = IsFunTy(LinT1(SrcLin_RmvMdy(A)))
'End Property
'
'Property Get SrcLin_IsTstSub(L$) As Boolean
'SrcLin_IsTstSub = True
'If IsPfx(L, "Sub Tst()") Then Exit Property
'If IsPfx(L, "Sub Tst()") Then Exit Property
'If IsPfx(L, "Friend Sub Tst()") Then Exit Property
'If IsPfx(L, "Sub ZZ__Tst()") Then Exit Property
'If IsPfx(L, "Sub ZZ__Tst()") Then Exit Property
'If IsPfx(L, "Friend Sub ZZ__Tst()") Then Exit Property
'SrcLin_IsTstSub = False
'End Property
'
'Property Get SrcLin_Mdy$(L)
'Dim A$
'A = "Private": If IsPfx(L, A) Then SrcLin_Mdy = A: Exit Property
'A = "Public":  If IsPfx(L, A) Then SrcLin_Mdy = A: Exit Property
'A = "Friend":  If IsPfx(L, A) Then SrcLin_Mdy = A: Exit Property
'End Property
'
'Property Get SrcLin_MthNm$(A)
'Dim L$: L = SrcLin_RmvMdy(A)
'Dim B$: B = LinShiftMthTy(L)
'If B = "" Then Exit Property
'SrcLin_MthNm = LinNm(L)
'End Property
'
'Property Get SrcLin_RmvMdy$(L)
'Dim A$
'A = "": If IsPfx(L, A) Then SrcLin_RmvMdy = RmvPfx(L, A): Exit Property
'A = "Public ":  If IsPfx(L, A) Then SrcLin_RmvMdy = RmvPfx(L, A): Exit Property
'A = "Friend ":  If IsPfx(L, A) Then SrcLin_RmvMdy = RmvPfx(L, A): Exit Property
'SrcLin_RmvMdy = L
'End Property
'
'Property Get SrcAllMthFmLnoAy(A$()) As Integer()
'Dim J%, O%()
'For J = 0 To UB(A)
'    If SrcLin_IsMth(A(J)) Then
'        Push O, J + 1 ' Return as Lno not index, it is J+1, not J
'    End If
'Next
'SrcAllMthFmLnoAy = O
'End Property
'
'Property Get SrcAllMthLinAy(A$()) As String()
'Dim L%(): L = SrcAllMthFmLnoAy(A)
'Dim O$(), LL
'For Each LL In L
'    Push O, SrcContLin(A, LL - 1)
'Next
'SrcAllMthLinAy = O
'End Property
'
'Property Get SrcContLin$(A$(), Lno%)
'Dim O$(), J%, L$
'For J = Lno To Sz(A)
'    L = A(J - 1) 'J-1 is Lno to Lx
'    If Right(L, 2) <> " _" Then
'        Push O, L
'        SrcContLin = Join(O, "")
'        Exit Property
'    End If
'    Push O, RmvLasChr(L)
'Next
'ErImposs
'End Property
'
'Property Get SrcDclLinCnt%(A$())
'Dim I&
'    I = SrcFstMthLx(A)
'    If I = -1 Then
'        SrcDclLinCnt = Sz(A)
'        Exit Property
'    End If
'    I = SrcMthRmkLx(A, I)
'Dim O&, L$
'    For I = I - 1 To 0 Step -1
'        If SrcLin_IsCd(A(I)) Then
'            O = I + 1
'            GoTo X
'        End If
'    Next
'X:
'SrcDclLinCnt = O
'End Property
'
'Property Get SrcDclLines$(A$())
'SrcDclLines = Join(SrcDclLy(A), vbCrLf)
'End Property
'
'Property Get SrcDclLy(A$()) As String()
'If Sz(A) = 0 Then Exit Property
'Dim N&
'   N = SrcDclLinCnt(A)
'If N <= 0 Then Exit Property
'SrcDclLy = LyTrimEnd(AyFstNEle(A, N))
'End Property
'
'Property Get Src_Dic_Of_MthKey_MthLines(A$(), Optional PjNm$, Optional MdNm$) As Dictionary
'Dim N$(): N = SrcMthNy(A)
'If Sz(N) = 0 Then Exit Property
'Dim O As New Dictionary
'    Dim Nm
'    Dim L$, K$, Lines$
'    For Each Nm In N
'        L = SrcMthLin(A, Nm)
'        K = MthLin_MthKey(L, PjNm, MdNm)
'        Lines = Src_MthLines_ByNm(A, Nm)
'        O.Add K, Lines
'    Next
'Set Src_Dic_Of_MthKey_MthLines = O
'End Property
'Property Get DicS1S2Ay(A As Dictionary) As S1S2()
'Dim O() As S1S2, K
'For Each K In A.Keys
'    PushObj O, S1S2(K, A(K))
'Next
'DicS1S2Ay = O
'End Property
'Sub DicBrw(A As Dictionary)
'S1S2Ay_Brw DicS1S2Ay(A)
'End Sub
'Property Get Src_Dic_Of_MthNm_MthLines(A$()) As Dictionary
'Dim O As Dictionary:
'If Sz(A) = 0 Then
'    Set O = New Dictionary
'    O.Add FmtQQ("*Empty Md"), ""
'    Set Src_Dic_Of_MthNm_MthLines = O
'    Exit Property
'End If
'Dim N
'For Each N In SrcMthNy(A)
'    O.Add N, Src_MthLines_ByNm(A, N)
'Next
'Dim D$: D = SrcDclLines(A)
'    If D <> "" Then O.Add "*Dcl", D
'
'Set Src_Dic_Of_MthNm_MthLines = O
'End Property
'
'Property Get SrcEndLx(A$(), MthLx)
'Dim F$: F = "End " & LinFunTy(A(MthLx))
'Dim J%
'For J = MthLx + 1 To UB(A)
'    If IsPfx(A(J), F) Then SrcEndLx = J: Exit Property
'Next
'End Property
'
'Property Get SrcFstMthLx&(A$())
'Dim J%
'For J = 0 To UB(A)
'   If SrcLin_IsMth(A(J)) Then
'       SrcFstMthLx = J
'       Exit Property
'   End If
'Next
'SrcFstMthLx = -1
'End Property
'
'Property Get SrcMthBdyFmToLno(A$(), MthNm$) As FmToLno
'Dim P As FmToLno
'    P = SrcMthFmToLno(A, MthNm)
'Dim FmLno%, Fnd As Boolean
'For FmLno = P.FmLno To P.ToLno
'    If Not LasChr(A(FmLno - 1)) = "_" Then
'        FmLno = FmLno + 1
'        Fnd = True
'        Exit For
'    End If
'Next
'If Not Fnd Then Stop
'With SrcMthBdyFmToLno
'    .FmLno = FmLno
'    .ToLno = P.ToLno - 1
'End With
'End Property
'
'Property Get Src_MthLines_ByMthFmLx$(A$(), MthFmLx)
'Dim P1$
'    P1 = SrcMthRmkLines(A, MthFmLx)
'Dim P2$
'    Dim L2%
'    L2 = SrcEndLx(A, MthFmLx)
'    P2 = Join(AyWhFmTo(A, MthFmLx, L2), vbCrLf)
'If P1 = "" Then
'    Src_MthLines_ByMthFmLx = P2
'Else
'    Src_MthLines_ByMthFmLx = P1 & vbCrLf & P2
'End If
'End Property
'
'Property Get SrcMthFmLno%(A$(), MthNm, Optional FmIx% = 0)
'Dim J%
'For J = FmIx To UB(A)
'    If SrcLin_MthNm(A(J)) = MthNm Then
'        SrcMthFmLno = J + 1 ' Return as Lno not index, it is J+1, not J
'        Exit Property
'    End If
'Next
'End Property
'
'Property Get SrcMthFmLnoAy(A$(), MthNm) As Integer()
'Dim L%
'L = SrcMthFmLno(A, MthNm): If L <= 0 Then Exit Property
'Dim O%(): Push O, L
'Dim S$: S = A(L - 1) ' SrcLin
'If SrcLin_FunTy(S) = "Property" Then
'    L = SrcMthFmLno(A, MthNm, L)
'    If L > 0 Then Push O, L
'End If
'SrcMthFmLnoAy = O
'End Property
'
'Property Get SrcMthFmToLno(A$(), MthNm$) As FmToLno
'If Sz(A) = 0 Then Exit Property
'Dim F%, T%
'F = SrcMthFmLno(A, MthNm)
'T = SrcMthToLno(A, F)
'With SrcMthFmToLno
'    .FmLno = F
'    .ToLno = T
'End With
'End Property
'
'Property Get SrcMthKy(A$(), Optional PjNm$ = "Pj", Optional MdNm$ = "Md", Optional IsSngLinFmt As Boolean) As String()
'Dim L%(): L = SrcAllMthFmLnoAy(A)
'If Sz(L) = 0 Then Exit Property
'Dim O$()
'    Dim MthLno
'    For Each MthLno In L
'        Push O, MthLin_MthKey(A(MthLno - 1), PjNm, MdNm, IsSngLinFmt)
'    Next
'SrcMthKy = O
'End Property
'
'Property Get SrcMthLin$(A$(), MthNm)
'Dim L%: L = SrcMthFmLno(A, MthNm)
'SrcMthLin = SrcContLin(A, L)
'End Property
'
'Property Get Src_MthLines_ByNm$(A$(), MthNm)
'Dim L%(): L = SrcMthFmLnoAy(A, MthNm)
'If Sz(L) = 0 Then Exit Property
'Dim MthLno, O$()
'For Each MthLno In L
'    Push O, Src_MthLines_ByMthFmLx(A, MthLno - 1)
'Next
'Src_MthLines_ByNm = Join(O, vbCrLf & vbCrLf)
'End Property
'Property Get Mdy0_MdyAy(A$) As String()
'
'End Property
'Property Get SrcMthNy(A$(), Optional MthNmPatn$ = ".", Optional IsNoSrt As Boolean, Optional Mdy0$) As String()
'Dim L%(): L = SrcAllMthFmLnoAy(A)
'If Sz(L) = 0 Then Exit Property
'Dim O$()
'    Dim MdyAy$(): MdyAy = Mdy0_MdyAy(Mdy0)
'    Dim MthLno, Lin$, N$, R As RegExp, M$
'    Set R = Re(MthNmPatn)
'    For Each MthLno In L
'        Lin = A(MthLno - 1)
'        N = MthLin_MthNm(Lin)
'        If R.Test(N) Then
'            M = SrcLin_Mdy(Lin)
'            If Mdy_IsSel(M, MdyAy) Then
'                PushNoDup O, N
'            End If
'        End If
'    Next
'If IsNoSrt Then
'    SrcMthNy = O
'Else
'    SrcMthNy = AySrt(O)
'End If
'End Property
'
'Property Get SrcMthRmkLines$(A$(), MthLx)
'Dim O$(), J%, L$, I%
'Dim Lx&: Lx = SrcMthRmkLx(A, MthLx)
'
'For J = Lx To MthLx - 1
'    L = Trim(A(J))
'    If L = "" Or L = "'" Then
'    ElseIf Left(L, 1) = "'" Then
'        Push O, L
'    Else
'         'Er in SrcMthRmkLx
'        Stop
'    End If
'Next
'SrcMthRmkLines = Join(O, vbCrLf)
'End Property
'
'Property Get SrcMthRmkLx&(A$(), MthLx)
'Dim M1&
'    Dim J&
'    For J = MthLx - 1 To 0 Step -1
'        If SrcLin_IsCd(A(J)) Then
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
'SrcMthRmkLx = M2
'End Property
'
'Property Get SrcMthToLno%(A$(), FmLno%)
'Dim T$: T = LinFunTy(A(FmLno - 1))
'If T = "" Then Stop
'Dim B$: B = "End " & T
'Dim J%
'For J = FmLno To UB(A)
'    If IsPfx(A(J), B) Then
'        SrcMthToLno = J + 1
'        Exit Property
'    End If
'Next
'Stop
'End Property
'
'Property Get SrcSrtRpt(A$(), PjNm$, MdNm$) As DCRslt
'Dim B$(): B = SrcSrtedLy(A)
'Dim A1 As Dictionary
'Dim B1 As Dictionary
'Set A1 = Src_Dic_Of_MthKey_MthLines(A, PjNm, MdNm)
'Set B1 = Src_Dic_Of_MthKey_MthLines(B, PjNm, MdNm)
'SrcSrtRpt = DicCmp(A1, B1)
'End Property
'
'Property Get SrcSrtRptLy(A$(), PjNm$, MdNm$) As String()
'SrcSrtRptLy = DCRslt_Ly(SrcSrtRpt(A, PjNm, MdNm))
'End Property
'
'Property Get SrcSrtedBdyLines$(A$())
'If Sz(A) = 0 Then Exit Property
'Dim D As Dictionary
'Dim D1 As Dictionary
'    Set D = Src_Dic_Of_MthKey_MthLines(A)
'    Set D1 = DicSrt(D)
'Dim O$()
'Dim K
'   For Each K In D1.Keys
'       Push O, vbCrLf & D1(K)
'   Next
'SrcSrtedBdyLines = JnCrLf(O)
'End Property
'Property Get DicSrt(A As Dictionary) As Dictionary
'Dim Ky(): Ky = A.Keys
'If Sz(Ky) = 0 Then Set DicSrt = New Dictionary: Exit Property
'Dim Ky1(): Ky1 = AySrt(Ky)
'Dim O As New Dictionary
'Dim K
'For Each K In Ky1
'    O.Add K, A(K)
'Next
'Set DicSrt = O
'End Property
'
'Property Get SrcSrtedLines$(A$())
'Dim O$(), A1$, A2$, A3$, A4$
'A1 = SrcDclLines(A)
'A2 = LinesTrimEnd(SrcDclLines(A))
'A3 = SrcSrtedBdyLines(A)
'A4 = LasChr(A3)
'If A4 = vbCr Or A4 = vbLf Then Stop
'PushNonEmp O, A2
'PushNonEmp O, A3
'SrcSrtedLines = Join(O, vbCrLf)
'End Property
'
'Property Get SrcSrtedLy(A$()) As String()
'SrcSrtedLy = SplitLines(SrcSrtedLines(A))
'End Property
'
'Property Get SslSy(Ssl) As String()
'SslSy = Split(Trim(RplDblSpc(Ssl)), " ")
'End Property
'
'Property Get StrAlignL$(S$, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
'Const CSub$ = "StrAlignL"
'Dim L%: L = Len(S)
'If L > W Then
'    If ErIfNotEnoughWdt Then
'        Stop
'        'Er CSub, "Len({S)) > {W}", S, W
'    End If
'    If DoNotCut Then
'        StrAlignL = S
'        Exit Property
'    End If
'End If
'
'If W >= L Then
'    StrAlignL = S & Space(W - L)
'    Exit Property
'End If
'If W > 2 Then
'    StrAlignL = Left(S, W - 2) + ".."
'    Exit Property
'End If
'StrAlignL = Left(S, W)
'End Property
'
'Sub StrBrw(A$)
'Dim T$:
'T = TmpFt
'StrWrt A, T
'Shell FmtQQ("code.cmd ""?""", T), vbMaximizedFocus
''Shell FmtQQ("notepad.exe ""?""", T), vbMaximizedFocus
'End Sub
'
'Property Get StrDup$(S, N%)
'Dim O$, J%
'For J = 0 To N - 1
'    O = O & S
'Next
'StrDup = O
'End Property
'
'Property Get StrNy(A) As String()
'Dim O$: O = RplPun(A)
'Dim O1$(): O1 = AyUniqAy(SslSy(O))
'Dim O2$()
'Dim J%
'For J = 0 To UB(O1)
'    If Not IsDigit(FstChr(O1(J))) Then Push O2, O1(J)
'Next
'StrNy = O2
'End Property
'
'Sub StrWrt(A, Ft$, Optional IsNotOvrWrt As Boolean)
'Fso.CreateTextFile(Ft, Overwrite:=Not IsNotOvrWrt).Write A
'End Sub
'
'Property Get Sz&(Ay)
'On Error Resume Next
'Sz = UBound(Ay) + 1
'End Property
'
'Property Get TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
'Dim Fnn$
'If Fnn0 = "" Then
'    Fnn = TmpNm
'Else
'    Fnn = Fnn0
'End If
'TmpFfn = TmpPth(Fdr) & Fnn & Ext
'End Property
'
'Property Get TmpFt$(Optional Fdr$, Optional Fnn$)
'TmpFt = TmpFfn(".txt", Fdr, Fnn)
'End Property
'
'Property Get TmpNm$()
'Static X&
'TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
'X = X + 1
'End Property
'
'Property Get TmpPth$(Optional Fdr$)
'Dim X$
'   If Fdr <> "" Then
'       X = Fdr & "\"
'   End If
'Dim O$
'   O = TmpPthHom & X:   PthEns O
'   O = O & TmpNm & "\": PthEns O
'   PthEns O
'TmpPth = O
'End Property
'
'Property Get TmpPthHom$()
'Static X$
'If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
'TmpPthHom = X
'End Property
'
'Property Get VarStr$(A)
'If IsPrim(A) Then VarStr = A: Exit Property
'If IsNothing(A) Then VarStr = "#Nothing": Exit Property
'If IsEmpty(A) Then VarStr = "#Empty": Exit Property
'If IsObject(A) Then
'    Dim T$
'    T = TypeName(A)
'    Select Case T
'    Case "CodeModule"
'        Dim M As CodeModule
'        Set M = A
'        VarStr = FmtQQ("*Md{?}", M.Parent.Name)
'        Exit Property
'    End Select
'    VarStr = "*" & T
'    Exit Property
'End If
'
'If IsArray(A) Then
'    Dim Ay: Ay = A: ReDim Ay(0)
'    T = TypeName(Ay(0))
'    VarStr = "*[" & T & "]"
'    Exit Property
'End If
'Stop
'End Property
'
'Property Get UB&(Ay)
'UB = Sz(Ay) - 1
'End Property
'
'Property Get VbeDupMthDrs(A As Vbe, Optional IsNoSrt As Boolean, Optional ExclPjNy0, Optional IsSamMthBdyOnly As Boolean) As Drs
'Dim Fny$(), Dry()
'Fny = SplitSsl("Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src")
'Dry = VbeDupMthDry(A, ExclPjNy0:=ExclPjNy0, IsSamMthBdyOnly:=IsSamMthBdyOnly)
'Set VbeDupMthDrs = Drs(Fny, Dry)
'End Property
'
'Property Get VbeDupMthDry(A As Vbe, Optional IsNoSrt As Boolean, Optional ExclPjNy0, Optional IsSamMthBdyOnly As Boolean) As Variant()
'Dim N$(): N = VbeMthFNy(A, IsNoSrt:=IsNoSrt, ExclPjNy0:=ExclPjNy0)
'Dim N1$(): N1 = MthFNy_DupMthFNy(N)
'    If IsSamMthBdyOnly Then
'        N1 = DupMthFNy_SamMthBdyMthFNy(N1, A)
'    End If
'Dim GpAy()
'    GpAy = DupMthFNy_GpAy(N1)
'    If Sz(GpAy) = 0 Then Exit Property
'Dim O()
'    Dim Gp
'    For Each Gp In GpAy
'        PushAy O, DupMthFNyGp_Dry(CvSy(Gp))
'    Next
'VbeDupMthDry = O
'End Property
'
'Property Get VbeDupMthFNy(A As Vbe, Optional IsNoSrt As Boolean, Optional ExclPjNy0, Optional IsSamMthBdyOnly As Boolean) As String()
'Dim N$(): N = VbeMthFNy(A, IsNoSrt:=IsNoSrt, ExclPjNy0:=ExclPjNy0)
'Dim N1$(): N1 = MthFNy_DupMthFNy(N)
'If IsSamMthBdyOnly Then
'    N1 = DupMthFNy_SamMthBdyMthFNy(N1, A)
'End If
'VbeDupMthFNy = N1
'End Property
'
'Property Get VbeDupFunLy(A As Vbe) As String()
'Dim I, O As New Dictionary
'For Each I In VbePjAy(A)
'    Set O = DupFunDic_Add(O, PjFunBdyDic(CvPj(I)))
'Next
'VbeDupFunLy = DupFunDic_Ly(O)
'End Property
'
'Property Get VbeDupMdNy(A As Vbe) As String()
'Dim O$()
'Stop '
'VbeDupMdNy = O
'End Property
'
'Sub VbeExport(A As Vbe)
'OyDo VbePjAy(A), "PjExport"
'End Sub
'
'Property Get VbeMthFNy(A As Vbe, Optional IsNoSrt As Boolean, Optional ExclPjNy0) As String()
'Dim Ay() As VBProject
'    Ay = VbePjAy(A, ExclPjNy0:=ExclPjNy0)
'If Sz(Ay) = 0 Then Exit Property
'Dim O$(), I
'For Each I In Ay
'    PushAy O, PjMthFNy(CvPj(I), IsNoSrt:=True)
'Next
'If IsNoSrt Then
'    VbeMthFNy = O
'Else
'    VbeMthFNy = AySrt(O)
'End If
'End Property
'
'Property Get VbeFstQPj(A As Vbe) As VBProject
'Dim I
'For Each I In A.VBProjects
'    If FstChr(CvPj(I).Name) = "Q" Then
'        Set VbeFstQPj = I
'        Exit Property
'    End If
'Next
'End Property
'
'Property Get VbeMdPjNy(A As Vbe, MdNm$) As String()
'Dim I, O$()
'For Each I In VbePjAy(A)
'    If PjHasCmp(CvPj(I), MdNm) Then
'        Push O, CvPj(I).Name
'    End If
'Next
'VbeMdPjNy = O
'End Property
'
'Property Get VbeMthKy(A As Vbe, Optional IsSngLinFmt As Boolean) As String()
'Dim O$(), I
'For Each I In VbePjAy(A)
'    PushAy O, PjMthKy(CvPj(I), IsSngLinFmt)
'Next
'VbeMthKy = O
'End Property
'
'Property Get VbeMthNy(A As Vbe, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Mdy$) As String()
'Dim Ay() As VBProject: Ay = VbePjAy(A, MdNmPatn)
'If Sz(Ay) = 0 Then Exit Property
'Dim I, O$()
'For Each I In Ay
'    PushAy O, PjMthNy(CvPj(I), MthNmPatn, MdNmPatn, Mdy)
'Next
'VbeMthNy = O
'End Property
'
'Property Get VbeMthNy_OfInproper(A As Vbe) As String()
'Dim I, O$()
'For Each I In VbePjAy(A)
'    PushAy O, PjMthNy_OfInproper(CvPj(I))
'Next
'VbeMthNy_OfInproper = O
'End Property
'
'Property Get VbePjAy(A As Vbe, Optional MdNmPatn$ = ".", Optional ExclPjNy0) As VBProject()
'Dim I, O() As VBProject
'Dim R As RegExp
'Set R = Re(MdNmPatn)
'Dim N$()
'Dim Nm$
'Dim X As Boolean
'    N = DftNy(ExclPjNy0)
'    X = Sz(N) > 0
'For Each I In A.VBProjects
'    Nm = CvPj(I).Name
'    If X Then
'        If AyHas(N, Nm) Then GoTo XX
'    End If
'    If R.Test(Nm) Then
'        PushObj O, I
'    End If
'XX:
'Next
'VbePjAy = O
'End Property
'
'Property Get VbePjNy(A As Vbe) As String()
'VbePjNy = ItrNy(A.VBProjects)
'End Property
'
'Property Get VbeSrcPth(A As Vbe)
'Dim Pj As VBProject:
'Set Pj = VbeFstQPj(A)
'Dim Ffn$: Ffn = PjFfn(Pj)
'If Ffn = "" Then Exit Property
'VbeSrcPth = FfnPth(Pj.Filename)
'End Property
'
'Sub VbeSrcPthBrw(A As Vbe)
'PthBrw VbeSrcPth(A)
'End Sub
'
'Sub VbeSrt(A As Vbe)
'Dim I
'For Each I In VbePjAy(A)
'    PjSrt CvPj(I)
'Next
'End Sub
'
'Property Get VbeSrtRptLy(A As Vbe) As String()
'Dim Ay() As VBProject: Ay = VbePjAy(A)
'Dim O$(), I, M As VBProject
'For Each I In Ay
'    Set M = I
'    PushAy O, PjSrtRptLy(M)
'Next
'VbeSrtRptLy = O
'End Property
'
'Property Get WbAddWs(A As Workbook, Optional WsNm$ = "Sheet1") As Worksheet
'Dim O As Worksheet
'Set O = A.Sheets.Add(A.Sheets(1))
'If WsNm <> "" Then
'   O.Name = WsNm
'End If
'Set WbAddWs = O
'End Property
'
'Property Get WsA1(A As Worksheet) As Range
'Set WsA1 = A.Cells(1, 1)
'End Property
'
'Property Get WsRC(A As Worksheet, R, C) As Range
'Set WsRC = A.Cells(R, C)
'End Property
'
'Sub WsVis(A As Worksheet)
'A.Application.Visible = True
'End Sub
'
'Property Get Xls() As Excel.Application
'Static Y As Excel.Application
'On Error GoTo X
'Dim A$: A = Y.Name
'Set Xls = Y
'Exit Property
'X:
'Set Y = New Excel.Application
'Set Xls = Y
'End Property
'
'Property Get XlsHasAddInFn(A As Excel.Application, AddInFn) As Boolean
'Dim I As Excel.AddIn
'Dim N$: N = UCase(AddInFn)
'For Each I In A.AddIns
'    If UCase(I.Name) = N Then XlsHasAddInFn = True: Exit Property
'Next
'End Property
'
'
'Private Sub ZZ_Pj()
'Ass "QAcs" = Pj("QAcs").Name
'End Sub
'
'Private Sub ZZ_Pj_Dic_Of_MthKey_MthLines()
'DicBrw Pj_Dic_Of_MthKey_MthLines(Pj("QVb"))
'End Sub
'
'Private Sub ZZ_PjRfLy()
'AyBrw PjRfLy(CurPj)
'End Sub
'
'Private Sub ZZ_PjSrtRptLy()
'AyBrw PjSrtRptLy(Pj("QSqTp"))
'End Sub
'
'Private Sub ZZ_Pj_TstClass_Bdy()
'Debug.Print Pj_TstClass_Bdy(Pj("QVb"))
'End Sub
'
'Private Sub ZZ_ReRpl()
'Dim R As RegExp: Set R = Re("(.+)(m[ae]n)(.+)")
'Dim Act$: Act = R.Replace("a men is male", "$1male$3")
'Ass Act = "a male is male"
'End Sub
'
'Private Sub ZZ_S1S2Ay_FmtLy()
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
'Act = S1S2Ay_FmtLy(A)
'AyBrw Act
'Exit Sub
'XX:
'    Set A(I) = S1S2(RplVBar(A1), RplVBar(A2))
'    Return
'End Sub
'
'Private Sub ZZ_SrcDclLinCnt()
'Dim B$(), A%
'
'B = ZZSrc
'A = SrcDclLinCnt(B)
'Ass A = 43
'
'B = MdSrc(Md("QSqTp.SqTp2"))
'A = SrcDclLinCnt(B)
'Ass A = 688
'End Sub
'
'Private Sub ZZ_SrcDclLines()
'Const P$ = "QSqTp"
'Const M$ = "SalRpt__CrdTyLvs_CrdExpr__Tst"
'Dim Md As CodeModule: Set Md = CurVbe.VBProjects(P).VBComponents(M).CodeModule
'Dim A$(): A = MdSrc(Md)
'Stop
'Dim B$: B = SrcDclLines(A)
'Stop
'StrBrw B
'End Sub
'
'Private Sub ZZ_SrcMthFmLnoAy()
'AyDmp SrcMthFmLnoAy(ZZSrc, "ZZA")
'End Sub
'
'Private Sub ZZ_Md_Dic_Of_MthKey_MthLines()
'Dim A As Dictionary: Set A = Md_Dic_Of_MthKey_MthLines(Md("QVb.M_Ay"))
'DicBrw A
'End Sub
'
'Private Sub ZZ_SrcSrtRptLy()
'AyBrw SrcSrtRptLy(ZZSrc, "QTool", "F_Tool")
'End Sub
'
'Private Sub ZZ_SrcSrtedBdyLines()
'StrBrw SrcSrtedBdyLines(ZZSrc)
'End Sub
'
'Private Sub ZZ_SrcSrtedLines()
'StrBrw SrcSrtedLines(ZZSrc)
'End Sub
'
'Private Sub ZZ_SrcSrtedLy()
'AyBrw SrcSrtedLy(ZZSrc)
'End Sub
'
'Private Sub ZZ_StrNy()
'Dim S$: S = MdLines(CurMd)
'AyBrw AySrt(StrNy(S))
'End Sub
'
'Private Sub ZZ_VbeMthNy()
'AyBrw VbeMthNy(CurVbe)
'End Sub
'
'
'Private Property Get ZZMd() As CodeModule
'Set ZZMd = Md("QTool.G_Tool")
'End Property
'
'Private Property Get ZZMth(MthNm$) As Mth
'Set ZZMth = Mth(ZZMd, MthNm)
'End Property
'
'Private Property Get ZZSrc() As String()
'ZZSrc = MdSrc(Md("F_Tool"))
'End Property
'
'Private Sub ZZ_Add_ZZA_Property()
'Dim S$
'S = "Private Property Get ZZA()|End Property||Property Set ZZA(A)|End Property"
'S = Replace(S, "|", vbCrLf)
'With CurMd
'    .InsertLines .CountOfLines + 1, S
'End With
'End Sub
'
'Private Sub ZZ_Dcl_BefAndAft_Srt()
'Const MdNm$ = "VbStrRe"
'Dim A$() ' Src
'Dim B$() ' Src->Srt
'Dim A1$ 'Src->Dcl
'Dim B1$ 'Src->Src->Dcl
'A = MdSrc(Md("QSqTp.SalRpt"))
'B = SrcSrtedLy(A)
'A1 = SrcDclLines(A)
'B1 = SrcDclLines(B)
'If A1 <> B1 Then Stop
'End Sub
'
'Private Sub ZZ_Shw_Mth()
'Shw_Mth "QTool.F_Tool.DDN_BrkAsg"
'End Sub
'
'Private Sub ZZ_PjSrtRptWb()
'Dim O As Workbook: Set O = PjSrtRptWb(CurPj, Vis:=True)
'Stop
'End Sub
'
'Private Sub ZZ_Pj_Compile()
'PjCompile Pj("QVb")
'End Sub
'
'Private Sub ZZ_ReMatch()
'Dim A As MatchCollection
'Dim R  As RegExp: Set R = Re("m[ae]n")
'Set A = R.Execute("alskdflfmensdklf")
'Stop
'End Sub
'
'Private Sub ZZ_Shw_Pj_SrtRptWb()
'Shw_Pj_SrtRptWb CurPj
'End Sub
'
'Private Sub ZZ_CurMdNm()
'Debug.Print CurMdNm
'End Sub
'
'Private Sub ZZ_CurVbe_PjNy()
'AyDmp CurVbe_PjNy
'End Sub
'
'Private Sub ZZ_MdAllMthLinAy()
'AyBrw MdAllMthLinAy(CurMd)
'End Sub
'
'Private Sub ZZ_MdGen_TstSub()
'MdGen_TstSub ZZMd
'End Sub
'
'Private Sub ZZ_MdMthNy()
'AyDmp MdMthNy(CurMd)
'End Sub
'
'Private Sub ZZ_MdRmv_TstSub()
'MdRmv_TstSub ZZMd
'End Sub
'
'Private Sub ZZ_MdSrtedLines()
'StrBrw MdSrtedLines(Md("QVb.M_Ay"))
'End Sub
'
'Private Sub ZZ_Md_TstSub_BdyLines()
'Debug.Print Md_TstSub_BdyLines(ZZMd)
'End Sub
'
'Private Sub ZZ_Md_TstSub_Lno()
'Debug.Print Md_TstSub_Lno(ZZMd)
'End Sub
'
'Private Sub ZZ_MthIsExist()
'Dim A As Mth: Set A = Mth(CurMd, "MthIsExist")
'Ass MthIsExist(A)
'End Sub
'
'Private Sub ZZ_MthLines()
'Debug.Print MthLines(Mth(CurMd, "ZZ_Mth_Lines"))
'End Sub
'
'Private Sub ZZ_MthLin()
'Debug.Print MthLin(ZZMth("ZZMth"))
'End Sub
'Private Sub MdAppVbl(A As CodeModule, Vbl)
'MdAppLines A, Replace(Vbl, "|", vbCrLf)
'End Sub
'Private Sub MdAppLines(A As CodeModule, Lines)
'A.InsertLines A.CountOfLines + 1, Lines
'End Sub
'Private Sub ZZ_MthRmv()
'Dim M As Mth: Set M = Mth(CurMd, "ZZA")
'MthRmv M
'Dim Lines$
'    Const C$ = "Property Get ZZA()|End Property|Property Let ZZA(A)|End Property"
'    Lines = Replace(C, "|", vbCrLf)
''AppLines
'    'MdAppVbl M.Md, Lines
'    M.Md.AddFromString Lines
''MthRmv M
'End Sub
'Property Get StrLin$(A)
'StrLin = A
'End Property
'Property Get AscIsLCase(A%) As Boolean
'If A < 97 Then Exit Property
'If A > 122 Then Exit Property
'AscIsLCase = True
'End Property
'Property Get AscIsUCase(A%) As Boolean
'If A < 65 Then Exit Property
'If A > 90 Then Exit Property
'AscIsUCase = True
'End Property
'Function FunNm_ProperMdNm$(A)
'Dim A0$
'    A0 = RmvPfx(A, "ZZ_")
'Dim P%
'Dim A1$
'    P = InStr(A0, "__")
'    If P > 0 Then
'        A1 = Left(A0, P - 1)
'    Else
'        A1 = A0
'    End If
'Dim P1%
'    P1 = InStr(A1, "_")
'
''--
'    If P1 > 0 Then
'        FunNm_ProperMdNm = Left(A1, P1 - 1)
'        Exit Function
'    End If
'Dim P2%
'Dim Fnd As Boolean
'    Dim C%
'    Fnd = False
'    For P2 = 2 To Len(A1)
'        C = Asc(Mid(A1, P2, 1))
'        If AscIsLCase(C) Then Fnd = True: Exit For
'    Next
''---
'    If Not Fnd Then
'        FunNm_ProperMdNm = A1
'        Exit Function
'    End If
'Dim P3%
'Fnd = False
'    For P3 = P2 + 1 To Len(A1)
'        C = Asc(Mid(A1, P3, 1))
'        If AscIsUCase(C) Then Fnd = True: Exit For
'    Next
''--
'If Fnd Then
'    FunNm_ProperMdNm = Left(A1, P3 - 1)
'    Exit Function
'End If
'FunNm_ProperMdNm = A1
'End Function
'
'Property Get MthProperMd(A As Mth) As CodeModule
''Mth here must be must belong to a StdMd
''Mth here must be Public, or,
''Mth name is ZZ_xxx, then it is ok to be private
'If Not MdIsStdMd(A.Md) Then Stop
'If Not IsPfx(A.Nm, "ZZ_") Then
'    If Not MthIsPub(A) Then Stop
'End If
'Dim Pj As VBProject
'Dim MdNm$
'    MdNm = "M_" & FunNm_ProperMdNm(A.Nm)
'    Set Pj = MdPj(A.Md)
'PjEns_Md Pj, MdNm
'Set MthProperMd = PjMd(Pj, MdNm)
'End Property
'
'Property Get AyMapS1S2Ay(A, MapFunNm$) As S1S2()
'Dim O() As S1S2, I
'If Sz(A) > 0 Then
'    For Each I In A
'        PushObj O, S1S2(I, Run(MapFunNm, I))
'    Next
'End If
'AyMapS1S2Ay = O
'End Property
'Sub AAA()
'AyBrw AyUniqAy(MdProperMdNy(CurMd))
'End Sub
'Sub ZZ_Md_FunNm_z_ProperMdNm_Brw()
'Md_FunNm_z_ProperMdNm_Brw CurMd
'End Sub
'Sub Md_FunNm_z_ProperMdNm_Brw(A As CodeModule)
'S1S2Ay_Brw Md_FunNm_z_ProperMdNm_S1S2Ay(A)
'End Sub
'Property Get MdProperMdNy(A As CodeModule) As String()
'Dim Ny$(): Ny = MdMthNy(A, IsNoMdNmPfx:=True, Mdy0:="Public")
'MdProperMdNy = AyMapSy(Ny, "FunNm_ProperMdNm")
'End Property
'Property Get Md_FunNm_z_ProperMdNm_S1S2Ay(A As CodeModule) As S1S2()
'Dim Ny$(): Ny = MdMthNy(A, IsNoMdNmPfx:=True, Mdy0:="Public")
'Md_FunNm_z_ProperMdNm_S1S2Ay = AyMapS1S2Ay(Ny, "FunNm_ProperMdNm")
'End Property
'Property Get MdIsStdMd(A As CodeModule) As Boolean
'MdIsStdMd = A.Parent.Type = vbext_ct_StdModule
'End Property
'
'
'
'
