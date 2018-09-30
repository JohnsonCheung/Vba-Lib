Attribute VB_Name = "AX"
Option Explicit
Type DCRslt
    Nm1 As String
    Nm2 As String
    AExcess As New Dictionary
    BExcess As New Dictionary
    ADif As New Dictionary
    BDif As New Dictionary
    Sam As New Dictionary
End Type
Public Fso As New FileSystemObject

Function DrsWhColEq(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, C)
Set DrsWhColEq = Drs(Fny, DryWhColEq(A.Dry, Ix, V))
End Function
Function DrsWhColGt(A As Drs, C$, V) As Drs
Dim Dry(), Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, C)
Set DrsWhColGt = Drs(Fny, DryWhColGt(A.Dry, Ix, V))
End Function
Function DrsAddValIdCol(A As Drs, ColNm$, Optional ColNmPfx$) As Drs
Dim Ix%, Fny$()
Fny = A.Fny
Ix = AyIx(Fny, ColNm): If Ix = -1 Then Stop
    Dim X$, Y$, C$
        C = ColNmPfx & ColNm
        X = C & "Id"
        Y = C & "Cnt"
    If AyHas(Fny, X) Then Stop
    If AyHas(Fny, Y) Then Stop
    PushIAy Fny, Array(X, Y)
Set DrsAddValIdCol = Drs(Fny, DryAddValIdCol(A.Dry, Ix))
End Function

Function DrsInsColBef(A As Drs, C$, FldNm$) As Drs
Set DrsInsColBef = DrsInsColXxx(A, C, FldNm, False)
End Function

Function DrsInsColAft(A As Drs, C$, FldNm$) As Drs
Set DrsInsColAft = DrsInsColXxx(A, C, FldNm, True)
End Function

Private Function DrsInsColXxx(A As Drs, C$, FldNm$, IsAft As Boolean) As Drs
Dim Fny1$(), Dry()
    Dim Ix&, Fny$()
    Fny = A.Fny
    Ix = AyIx(Fny, C): If Ix = -1 Then Stop
    If IsAft Then
        Ix = Ix + 1
    End If
    Fny1 = AyIns(Fny, FldNm, Ix)
    Dry = DryInsCol(A.Dry, C, Ix)
Set DrsInsColXxx = Drs(Fny1, Dry)
End Function

Private Sub ZZ_DrsGpFlat()
Dim Act As Drs, Drs2 As Drs, Drs1 As Drs, N1%, N2%
'Set Drs1 = VbeFun12Drs(CurVbe)
'N1 = Sz(Drs1.Dry)
'Set Drs2 = VbeMth12Drs(CurVbe)
'N2 = Sz(Drs2.Dry)
'Debug.Print N1, N2
Set Act = DrsGpFlat(Drs1, "Nm", "Lines")
DrsBrw Act
End Sub

Function CvAy(A) As Variant()
CvAy = A
End Function
Private Sub ZZ_DrsGpFlat_1()
Dim Act As Drs, D As Drs, Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Set D = Drs("A B C", CvAy(Array(Dr1, Dr2, Dr3)))
Set Act = DrsGpFlat(D, "A", "C")
Stop
DrsBrw Act
End Sub
Function DrsKeyCntDic(A As Drs, K$) As Dictionary
Dim Dry(), O As New Dictionary, Fny$(), Dr, Ix%, KK$
Fny = A.Fny
Ix = AyIx(Fny, K)
Dry = A.Dry
If Sz(Dry) > 0 Then
    For Each Dr In A.Dry
        KK = Dr(Ix)
        If O.Exists(KK) Then
            O(KK) = O(KK) + 1
        Else
            O.Add KK, 1
        End If
    Next
End If
Set DrsKeyCntDic = O
End Function

Function LinNPrm(A) As Byte
LinNPrm = SubStrCnt(BktStr(A), ",")
End Function

Private Sub Z_SubStrCnt()
Dim S$, SS$, A&, E&
S = "skfdj skldfskldf df "
SS = " "
E = 3
A = SubStrCnt(S, SS)
Ass A = E
End Sub

Function VbeHasPj(A As Vbe, PjNm) As Boolean
VbeHasPj = ItrHasNm(A.VBProjects, PjNm)
End Function
Function ItrHasNm(A, Nm) As Boolean
Dim I
For Each I In A
    If I.Name = Nm Then ItrHasNm = True: Exit Function
Next
End Function
Function ItrHasNmWhRe(A, Re As RegExp) As Boolean
Dim I
For Each I In A
    If Re.Test(I.Name) Then ItrHasNmWhRe = True: Exit Function
Next
End Function

Function TakBefOrAll$(S, Sep, Optional NoTrim As Boolean)
TakBefOrAll = Brk1(S, Sep, NoTrim).S1
End Function
Function TakAftOrAll$(S, Sep, Optional NoTrim As Boolean)
TakAftOrAll = Brk2(S, Sep, NoTrim).S2
End Function
Function TakAftMust$(A, Sep, Optional NoTrim As Boolean)
TakAftMust = Brk(A, Sep, NoTrim).S2
End Function
Function TakAft$(A, Sep, Optional NoTrim As Boolean)
TakAft = Brk1(A, Sep, NoTrim).S2
End Function
Function TakBef$(S, Sep$, Optional NoTrim As Boolean)
TakBef = Brk2(S, Sep, NoTrim).S1
End Function
Function TakBefMust$(S, Sep$, Optional NoTrim As Boolean)
TakBefMust = Brk(S, Sep, NoTrim).S1
End Function


Function AlignL$(A, W, Optional ErIFmnotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "AlignL"
If ErIFmnotEnoughWdt And DoNotCut Then
    Stop
    'Er CSub, "Both {ErIFmnotEnoughWdt} and {DontCut} cannot be True", ErIFmnotEnoughWdt, DoNotCut
End If
Dim S$: S = VarStr(A)
AlignL = StrAlignL(S, W, ErIFmnotEnoughWdt, DoNotCut)
End Function













Sub ZZ_AyIns()

End Sub













Function RunAv(FunNm$, Av())
Dim O
Select Case Sz(Av)
Case 0: O = Run(FunNm)
Case 1: O = Run(FunNm, Av(0))
Case 2: O = Run(FunNm, Av(0), Av(1))
Case 3: O = Run(FunNm, Av(0), Av(1), Av(2))
Case 4: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(FunNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case Else: Stop
End Select
RunAv = O
End Function




















Function CvNy(A) As String()
If IsStr(A) Then CvNy = SslSy(A): Exit Function
If IsSy(A) Then CvNy = A: Exit Function
Stop
End Function

Function IsInLikAy(A, LikAy) As Boolean
If Sz(LikAy) = 0 Then Exit Function
Dim Lik
For Each Lik In LikAy
    If A Like Lik Then IsInLikAy = True: Exit Function
Next
End Function

Function SqRg(A, At As Range) As Range
If Sz(A) = 0 Then Exit Function
Dim O As Range: Set O = CellReSz(At, A)
O.Value = A
Set SqRg = O
End Function

Function SqLo(A, At As Range, Optional LoNm$) As ListObject
Set SqLo = RgLo(SqRg(A, At), LoNm)
End Function

Function CellReSz(A As Range, Sq) As Range
Set CellReSz = RgRCRC(A, 1, 1, UBound(Sq, 1), UBound(Sq, 2))
End Function

Function CmpTyAyOf_Cls_and_Std() As vbext_ComponentType()
Dim O(1) As vbext_ComponentType
O(0) = vbext_ct_ClassModule
O(1) = vbext_ct_StdModule
CmpTyAyOf_Cls_and_Std = O
End Function

Function CmpTy_Nm$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ct_ClassModule: O = "*Cls"
Case vbext_ct_StdModule: O = "*Std"
Case vbext_ct_Document: O = "*Doc"
Case Else: Stop
End Select
CmpTy_Nm = O
End Function

Function CollAddPfx(A As Collection, Pfx) As Collection
Dim O As New Collection, I
For Each I In A
    O.Add Pfx & I
Next
Set CollAddPfx = O
End Function

Function CurXls() As Excel.Application
Set CurXls = Excel.Application
End Function
Function CurWb() As Workbook
Set CurWb = CurXls.ActiveWorkbook
End Function

Function CurWs() As Worksheet
Set CurWs = CurXls.ActiveSheet
End Function


Function CurCmp() As VBComponent
Set CurCmp = CurMd.Parent
End Function

Function CurFunDNm$()
Dim M$: M = CurMthNm
If M = "" Then Exit Function
If Not IsStdMd(CurMd) Then Exit Function
CurFunDNm = CurMdDNm & "." & M
End Function
Function CurSrc() As String()
CurSrc = MdSrc(CurMd)
End Function
Function CurMd() As CodeModule
Set CurMd = CurVbe.ActiveCodePane.CodeModule
End Function

Function CurMdDNm$()
CurMdDNm = MdDNm(CurMd)
End Function

Function CurMdNm$()
CurMdNm = CurCmp.Name
End Function

Function CurMth() As Mth
Dim Nm$: Nm = CurMthNm
If Nm = "" Then Stop
Set CurMth = Mth(CurMd, Nm)
End Function

Function CurMthDNm$()
CurMthDNm = CurMdDNm & "." & CurMthNm
End Function

Function CurMthNm$()
Dim L1&, L2&, C1&, C2&, K As vbext_ProcKind
Dim O$
With CurVbe.ActiveCodePane
    On Error GoTo X
    .GetSelection L1, C1, L2, C2
    On Error GoTo 0
    O = .CodeModule.ProcOfLine(L1, K)
End With
If O = "" Then Stop
CurMthNm = O
Exit Function
X:
End Function

Function CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Function

Function CurPjNm$()
CurPjNm = CurPj.Name
End Function

Function CurPjPth$()
CurPjPth = PjPth(CurPj)
End Function

Function CurVbe() As Vbe
Set CurVbe = CurXls.Vbe
End Function

Function CvFTNo(A) As FTNo
Set CvFTNo = A
End Function

Function CvFTIx(A) As FTIx
Set CvFTIx = A
End Function

Function CvMd(A) As CodeModule
Set CvMd = A
End Function
Function CvCmp(A) As VBComponent
Set CvCmp = A
End Function
Function CvS1S2(A) As S1S2
Set CvS1S2 = A
End Function
Function CvMth(A) As Mth
Set CvMth = A
End Function

Function CvPj(I) As VBProject
Set CvPj = I
End Function

Function CvSy(A) As String()
CvSy = A
End Function

Function DCRsltBrw(A As DCRslt)

End Function

Function DCRsltIsSam(A As DCRslt) As Boolean
With A
If .ADif.Count > 0 Then Exit Function
If .BDif.Count > 0 Then Exit Function
If .AExcess.Count > 0 Then Exit Function
If .BExcess.Count > 0 Then Exit Function
End With
DCRsltIsSam = True
End Function

Function DCRsltFmt(A As DCRslt) As String()
With A
Dim A1$(): A1 = DCRsltFmt__AExcess(.AExcess)
Dim A2$(): A2 = DCRsltFmt__BExcess(.BExcess)
Dim A3$(): A3 = DCRsltFmt__Dif(.ADif, .BDif)
Dim A4$(): A4 = DCRsltFmt__Sam(.Sam)
End With
DCRsltFmt = AyAddAp(A1, A2, A3, A4)
End Function

Function DCRsltFmt__AExcess(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, Ly$(), S1$, S2$, S(0) As S1S2
S2 = "!" & "Er AExcess"
For Each K In A.Keys
    S1 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2AyFmt(S)
    PushAy O, Ly
Next
DCRsltFmt__AExcess = O
End Function

Function DCRsltFmt__BExcess(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, Ly$(), S1$, S2$, S(0) As S1S2
S1 = "!" & "Er BExcess"
For Each K In A.Keys
    S2 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2AyFmt(S)
    PushAy O, Ly
Next
DCRsltFmt__BExcess = O
End Function

Function DCRsltFmt__Dif(A As Dictionary, B As Dictionary) As String()
If A.Count <> B.Count Then Stop
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S(0) As S1S2, Ly$()
For Each K In A
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & B(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2AyFmt(S)
    PushAy O, Ly
Next
DCRsltFmt__Dif = O
End Function

Function DCRsltFmt__Sam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S1S2
For Each K In A.Keys
    PushObj S, S1S2("*Same", K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K))
Next
DCRsltFmt__Sam = S1S2AyFmt(S)
End Function

Function DDNmThird$(A)
Dim Ay$(): Ay = Split(A, "."): If Sz(Ay) <> 3 Then Stop
DDNmThird = Ay(2)
End Function

Function DftMd(MdDNm0$)
If MdDNm0 = "" Then
    Set DftMd = CurMd
Else
    Set DftMd = Md(MdDNm0)
End If
End Function

Function DftMdDNm$(MdDNm0$)
If MdDNm0 = "" Then
    DftMdDNm = CurMdNm
Else
    DftMdDNm = MdDNm0
End If
End Function

Function DftMdySy(A$) As String()
DftMdySy = DftNy(A)
End Function

Function DftMth(MthDNm0$) As Mth
If MthDNm0 = "" Then
    Set DftMth = CurMth
    Exit Function
End If
Set DftMth = DMth(MthDNm0)
End Function

Function DftMthNm$(MthNm0$)
If MthNm0 = "" Then
    DftMthNm = CurMthNm
    Exit Function
End If
DftMthNm = MthNm0
End Function

Function DftNy(Ny0) As String()
Dim T As VbVarType: T = VarType(Ny0)
If T = vbEmpty Then Exit Function
If IsMissing(Ny0) Then Exit Function
If T = vbString Then
    DftNy = SplitSsl(Ny0)
    Exit Function
End If
DftNy = Ny0
End Function

Function DftPj(PjNm0$)
If PjNm0 = "" Then
    Set DftPj = CurPj
Else
    Set DftPj = Pj(PjNm0)
End If
End Function



Function DicAdd(A As Dictionary, B As Dictionary) As Dictionary
Dim O  As New Dictionary, I
For Each I In A.Keys
    O.Add I, A(I)
Next
For Each I In B.Keys
    O.Add I, B(I)
Next
Set DicAdd = O
End Function

Function DicClone(A As Dictionary) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, A(K)
Next
Set DicClone = O
End Function

Function DicCmp(A As Dictionary, B As Dictionary, Optional Nm1$ = "Fst", Optional Nm2$ = "Snd") As DCRslt
Dim O As DCRslt
Set O.AExcess = DicMinus(A, B)
Set O.BExcess = DicMinus(B, A)
Set O.Sam = DicAB_SamDic(A, B)
Dim DicAB(): DicAB = DicAB_SamKeyDifVal_DicPair(A, B)
    Set O.ADif = DicAB(0)
    Set O.BDif = DicAB(1)
O.Nm1 = Nm1
O.Nm2 = Nm2
DicCmp = O
End Function

Function DicHasAllKeyIsNm(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsNm(K) Then Exit Function
Next
DicHasAllKeyIsNm = True
End Function

Function DicHasAllValIsStr(A As Dictionary) As Boolean
Dim K
For Each K In A.Keys
    If Not IsStr(A(K)) Then Exit Function
Next
DicHasAllValIsStr = True
End Function

Function IsEqDic(A As Dictionary, B As Dictionary) As Boolean
Dim K(): K = A.Keys
If Sz(K) <> Sz(B.Keys) Then Exit Function
Dim KK, J%
For Each KK In K
    J = J + 1
    If IsEq(A(KK), B(KK)) Then Exit Function
Next
IsEqDic = True
Stop
End Function

Function DicMinus(A As Dictionary, B As Dictionary) As Dictionary
If A.Count = 0 Then Set DicMinus = New Dictionary: Exit Function
If B.Count = 0 Then Set DicMinus = DicClone(A): Exit Function
Dim O As New Dictionary, K
For Each K In A.Keys
   If Not B.Exists(K) Then O.Add K, A(K)
Next
Set DicMinus = O
End Function
Function DicS1S2Itr(A As Dictionary) As Collection
Dim O As New Collection, K
For Each K In A.Keys
    O.Add S1S2(K, A(K))
Next
Set DicS1S2Itr = O
End Function

Function DicS1S2Ay(A As Dictionary) As S1S2()
Dim O() As S1S2, K
For Each K In A.Keys
    PushObj O, S1S2(K, A(K))
Next
DicS1S2Ay = O
End Function

Function DicSrt(A As Dictionary) As Dictionary
Dim Ky(): Ky = A.Keys
If Sz(Ky) = 0 Then Set DicSrt = New Dictionary: Exit Function
Dim Ky1(): Ky1 = AySrt(Ky)
Dim O As New Dictionary
Dim K
For Each K In Ky1
    O.Add K, A(K)
Next
Set DicSrt = O
End Function

Function DicWb(A As Dictionary, Optional Vis As Boolean) As Workbook
'Assume each dic keys is name and each value is lines
'Prp-Wb is to create a new Wb with worksheet as the dic key and the lines are break to each cell of the sheet
Ass DicHasAllKeyIsNm(A)
Ass DicHasAllValIsStr(A)
Dim K, ThereIsSheet1 As Boolean
Dim O As Workbook: Set O = NewWb
Dim Ws As Worksheet
For Each K In A.Keys
    If K = "Sheet1" Then
        Set Ws = O.Sheets("Sheet1")
        ThereIsSheet1 = True
    Else
        Set Ws = O.Sheets.Add
        Ws.Name = K
    End If
    Ws.Range("A1").Value = LinesSqV(A(K))
Next
X: Set Ws = O
If Vis Then O.Application.Visible = True
End Function
Function DrsWs(A As Drs) As Worksheet
Dim O As Worksheet, R As Range
Set O = NewWs
Set R = AyRgH(A.Fny, WsA1(O))
If Sz(A.Dry) = 0 Then
    RgLo RgDecBtmR(R)
Else
    RgLo RgIncTopR(SqRg(DrySq(A.Dry), WsRC(O, 2, 1)))
End If
Set DrsWs = O
End Function
Function RgIncTopR(A As Range, Optional By% = 1) As Range
Set RgIncTopR = RgRR(A, 1 - By, A.Rows.Count)
End Function
Function RgDecBtmR(A As Range, Optional By% = 1) As Range
Set RgDecBtmR = RgRR(A, 1, A.Rows.Count + 1)
End Function
Function RgRR(A As Range, R1, R2) As Range
Set RgRR = RgRCRC(A, R1, 1, R2, A.Columns.Count)
End Function

Function DrsSq(A As Drs) As Variant()
Dim NCol&, NRow&, Dry(), Fny$()
    Fny = A.Fny
    Dry = A.Dry
    NCol = Max(DryNCol(Dry), Sz(Fny))
    NRow = Sz(Dry)
Dim O()
ReDim O(1 To 1 + NRow, 1 To NCol)
Dim C&, R&, Dr()
    For C = 1 To Sz(Fny)
        O(1, C) = Fny(C - 1)
    Next
    For R = 1 To NRow
        Dr = A(R - 1)
        For C = 1 To Min(Sz(Dr), NCol)
            O(R + 1, C) = Dr(C - 1)
        Next
    Next
DrsSq = O
End Function

Function DupMthFNyGpAy_AllSameCnt%(A)
If Sz(A) = 0 Then Exit Function
Dim O%, Gp
For Each Gp In A
    If DupMthFNyGp_IsDup(Gp) Then O = O + 1
Next
DupMthFNyGpAy_AllSameCnt = O
End Function

Function DupMthFNyGp_Dry(Ny$()) As Variant()
'Given Ny: Each Nm in Ny is FunNm:PjNm.MdNm
'          It has at least 2 ele
'          Each FunNm is same
'Return: N-Dr of Fields {Nm Mdy-1 Ty-1 Pj-1 Md-1 Mdy-2 Ty-2 Pj-2 Md-2 Src-1 Src-2 IsSam-Pj IsSam-Md IsSam-Src}
'        where N = Sz(Ny)-1
'        where each-field-(*-1)-of-Dr comes from Ny(0)
'        where each-field-(*-2)-of-Dr comes from Ny(1..)

Dim Md1$, Pj1$, Nm$
    FunFNm_BrkAsg Ny(0), Nm, Pj1, Md1
Dim Mth1 As Mth
    Set Mth1 = Mth(Md(Pj1 & "." & Md1), Nm)
Dim Src1$
    Src1 = MthLines(Mth1)
Dim Mdy1$, Ty1$
    MthBrkAsg Mth1, Mdy1, Ty1
Dim O()
    Dim J%
    For J = 1 To UB(Ny)
        Dim Pj2$, Nm2$, Md2$
            FunFNm_BrkAsg Ny(J), Nm2, Pj2, Md2: If Nm2 <> Nm Then Stop
        Dim Mth2 As Mth
            Set Mth2 = Mth(Md(Pj2 & "." & Md2), Nm)
            Dim Src2$
            Src2 = MthLines(Mth2)
        Dim Mdy2$, Ty2$
            MthBrkAsg Mth2, Mdy2, Ty2

        Push O, Array(Nm, _
                    Mdy1, Ty1, Pj1, Md1, _
                    Mdy2, Ty2, Pj2, Md2, Src1, Src2, Pj1 = Pj2, Md1 = Md2, Src1 = Src2)
    Next
DupMthFNyGp_Dry = O
End Function

Function DupMthFNyGp_IsDup(Ny) As Boolean
DupMthFNyGp_IsDup = AyIsAllEleEq(AyMap(Ny, "FunFNm_MthLines"))
End Function

Function DupMthFNy_GpAy(A$()) As Variant()
Dim O(), J%, M$()
Dim L$ ' LasMthNm
L = Brk(A(0), ":").S1
Push M, A(0)
Dim B As S1S2
For J = 1 To UB(A)
    Set B = Brk(A(J), ":")
    If L <> B.S1 Then
        Push O, M
        Erase M
        L = B.S1
    End If
    Push M, A(J)
Next
If Sz(M) > 0 Then
    Push O, M
End If
DupMthFNy_GpAy = O
End Function

Function EitherL(A) As Either
Asg A, EitherL.Left
EitherL.IsLeft = True
End Function

Function PjHasCmp(A As VBProject, Nm$) As Boolean
PjHasCmp = ItrHasNm(A.VBComponents, Nm)
End Function

Sub PjAddCmp(A As VBProject, Nm$, Ty As vbext_ComponentType)
If PjHasCmp(A, Nm) Then
    Debug.Print FmtQQ("PjAddCmp: Pj[?] already has Cmp[?]", A.Name, Nm)
    Exit Sub
End If
With A.VBComponents.Add(Ty)
    .Name = Nm
    .CodeModule.InsertLines 1, "Option Explicit"
    MdSav .CodeModule
End With
End Sub
Sub MdSav(A As CodeModule)

End Sub
Function EitherR(A) As Either
Asg A, EitherR.Right
End Function
Function EmpMdAy() As CodeModule
End Function
Function EmpAy() As Variant()
End Function

Function EmpIntAy() As Integer()
End Function

Function EmpRfAy() As Reference()
End Function
Function IsDic(A) As Boolean
IsDic = TypeName(A) = "Dictionary"
End Function
Function IsSyDic(A) As Boolean
Dim D As Dictionary, I, V
If Not IsDic(A) Then Exit Function
Set D = A
For Each I In D.Keys
    V = D(I)
    If Not IsSy(V) Then Exit Function
Next
IsSyDic = True
End Function

Function IsStrDic(A) As Boolean
Dim D As Dictionary, I
If Not IsDic(A) Then Exit Function
Set D = A
For Each I In D.Keys
    If Not IsStr(D(I)) Then Exit Function
Next
IsStrDic = True
End Function
Function DicAy_Mge(A() As Dictionary) As Dictionary
'Assume there is no duplicated key in each of the dic in A()
Dim O As New Dictionary
If Sz(A) > 0 Then
    Dim I
    For Each I In A
        DicPush O, CvDic(I)
    Next
End If
Set DicAy_Mge = O
End Function
Function CvDic(A) As Dictionary
Set CvDic = A
End Function
Sub DicPush(O As Dictionary, M As Dictionary)
'Assume there is no duplicated key
If M.Count = 0 Then Exit Sub
Dim K
For Each K In M.Keys
    O.Add K, M(K)
Next
End Sub
Function RmvUSfx$(A) ' Return upcase Sfx
Dim J%, Fnd As Boolean, C%
For J = Len(A) To 2 Step -1 ' don't find the first char if non-UCase, to use 'To 2'
    C = Asc(Mid(A, J, 1))
    If Not AscIsUCase(C) Then
        Fnd = True
        Exit For
    End If
Next
If Fnd Then
    RmvUSfx = Left(A, J)
Else
    RmvUSfx = A
End If
End Function
Function DicIsEmp(A As Dictionary) As Boolean
DicIsEmp = A.Count = 0
End Function

Function EmpDicAy() As Dictionary()
End Function
Function DicMap(A As Dictionary, ValMapFun$) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add K, Run(ValMapFun, A(K))
Next
Set DicMap = O
End Function
Function CvDicAy(A) As Dictionary()
CvDicAy = A
End Function
Function EmpSy() As String()
End Function






Function FTNoAyLinCnt%(A() As FTNo)
Dim O%, M
For Each M In A
    O = O + FTNoLinCnt(CvFTNo(M))
Next
End Function

Function FTNoLinCnt%(A As FTNo)
Dim O%
O = A.Tono - A.Fmno + 1
If O < 0 Then Stop
FTNoLinCnt = O
End Function

Function FTIxNo(A As FTIx) As FTNo
Set FTIxNo = FTNo(A.Fmix + 1, A.Toix + 1)
End Function

Function FTIxLinCnt%(A As FTIx)
Dim O%
O = A.Toix - A.Fmix + 1
If O < 0 Then Stop
FTIxLinCnt = O
End Function

Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim O$, L$, P&, I
L = Replace(QQVbl, "|", vbCrLf)
Dim Av(): Av = Ap
P = 1
For Each I In Av
    P = InStr(L, "?")
    If P = 0 Then FmtQQ = O & L: Exit Function
    O = O & Left(L, P - 1) & I
    L = Mid(L, P + 1)
Next
FmtQQ = O & L
End Function

Function MthKeyDrFny() As String()
MthKeyDrFny = SslSy("PjNm MdNm Priority Nm Ty Mdy")
End Function

Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function FunFNm_MdDNm$(A)
FunFNm_MdDNm = Brk(A, ":").S2
End Function

Function FunFNm_MthLines$(A)
FunFNm_MthLines = MthLines(MthFNm_Mth(A))
End Function


Function MthDryWhDup(MthDry(), Optional IsSamMthBdyOnly As Boolean) As String()
'MthFNy is in format of Pj Md ShtMdy ShtTy Nm Lines
If Sz(MthDry) = 0 Then Exit Function
Stop '
Dim Ny$(): Ny = DryStrCol(MthDry, 4)
Dim A1$(): A1 = AySrt(Ny)
Dim O$(), M$(), J&, Nm$
Dim L$ ' LasFunNm
L = MthFNm_Nm(A1(0))
Push M, A1(0)
For J = 1 To UB(A1)
    Nm = MthFNm_Nm(A1(J))
    If L = Nm Then
        Push M, A1(J)
    Else
        L = Nm
        If Sz(M) = 1 Then
            M(0) = A1(J)
        Else
            PushAy O, M
            Erase M
        End If
    End If
Next
If Sz(M) > 1 Then
    PushAy O, M
End If
MthDryWhDup = O
End Function

Private Sub Z_MthPfx()
Ass MthPfx("Add_Cls") = "Add"
End Sub
Function ItrPredSomTrue(A, Pred$) As Boolean
Dim I
For Each I In A
    If Run(Pred, I) Then ItrPredSomTrue = True: Exit Function
Next
End Function
Function ItrPredSomFalse(A, Pred$) As Boolean
Dim I
For Each I In A
    If Not Run(Pred, I) Then ItrPredSomFalse = True: Exit Function
Next
End Function
Function ItrPredAllTrue(A, Pred$) As Boolean
Dim I
For Each I In A
    If Not Run(Pred, I) Then Exit Function
Next
ItrPredAllTrue = True
End Function
Function ItrPredAllFalse(A, Pred$) As Boolean
Dim I
For Each I In A
    If Run(Pred, I) Then Exit Function
Next
ItrPredAllFalse = True
End Function
Private Sub ZZ_TimFun()
TimFun "ZZ_DicHasStrKy ZZ_DicHasStrKy1"
End Sub
Sub TimFun(FunNy0)
Dim B!, E!, F
For Each F In DftNy(FunNy0)
    B = Timer
    Run F
    E = Timer
    Debug.Print F, "<-- Run"; E - B
Next
End Sub
Private Sub ZZ_DicHasStrKy3()
TimFun "ZZ_DicHasStrKy ZZ_DicHasStrKy1"
End Sub
Private Sub ZZ_DicHasStrKy()
ZZ_DicHasStrKy__X "DicHasStrKy"
End Sub
Private Sub ZZ_DicHasStrKy1()
ZZ_DicHasStrKy__X "DicHasStrKy1"
End Sub

Private Sub ZZ_DicHasStrKy2()
Dim A As New Dictionary, Exp As Boolean, Act As Boolean
Dim J&
For J = 1 To 10000
    A.Add CStr(J), J
Next
Act = DicHasStrKy(A)
Exp = True
Ass Act = Exp

A.Add 10001, "X"
Act = DicHasStrKy(A)
Exp = False
Ass Act = Exp

End Sub
Private Sub ZZ_DicHasStrKy__X(X$)
Dim A As New Dictionary, Exp As Boolean, Act As Boolean
Dim J&
For J = 1 To 10000
    A.Add CStr(J), J
Next
Act = Run(X, A)
Exp = True
Ass Act = Exp

A.Add 10001, "X"
Act = Run(X, A)
Exp = False
Ass Act = Exp

End Sub
Function DicHasStrKy(A As Dictionary) As Boolean
DicHasStrKy = ItrPredAllTrue(A.Keys, "IsStr")
End Function
Function DicHasStrKy1(A As Dictionary) As Boolean
Dim I
For Each I In A.Keys
    If Not IsStr(I) Then Exit Function
Next
DicHasStrKy1 = True
End Function
Private Sub ZZ_MthPfx()
Dim Ay$(): Ay = VbeMthNy(CurVbe)
Dim Ay1$(): Ay1 = AyMapSy(Ay, "MthPfx")
WsVis AyabWs(Ay, Ay1)
End Sub
Private Sub ZZ_AyabWs()
Dim A, B
A = SslSy("A B C D E")
B = SslSy("1 2 3 4 5")
WsVis AyabWs(A, B)
Stop
End Sub

Private Sub Z_RmvPfxAy()
Dim A$, PfxAy$()
PfxAy = SslSy("ZZ_ Z_"): Ept = "ABC"
A = "Z_ABC": GoSub Tst
A = "ZZ_ABC": GoSub Tst
Exit Sub
Tst:
    Act = RmvPfxAy(A, PfxAy)
    C
    Return
End Sub
Function RmvPfxAyS$(A, PfxAy)
Dim P
For Each P In PfxAy
    If HasPfxS(A, P) Then
        RmvPfxAyS = Mid(A, Len(P) + 2)
        Exit Function
    End If
Next
RmvPfxAyS = A
End Function

Function RmvPfxAy$(A, PfxAy)
Dim P
For Each P In PfxAy
    If HasPfx(A, P) Then
        RmvPfxAy = RmvPfx(A, P)
        Exit Function
    End If
Next
RmvPfxAy = A
End Function

Sub Brw(A)
Select Case True
Case IsStr(A): StrBrw A
Case IsArray(A): AyBrw A
Case Else: Stop
End Select
End Sub
Sub ZZ_AyGpCntFmt()
Brw AyGpCntFmt(CurPjFunPfxAy)
End Sub

Function MthPfx$(MthNm$)
Dim A0$
    A0 = Brk1(RmvPfxAy(MthNm, SplitVBar("ZZ_|Z_")), "__").S1
With Brk2(A0, "_")
    If .S1 <> "" Then
        MthPfx = .S1
        Exit Function
    End If
End With
Dim P2%
Dim Fnd As Boolean
    Dim C%
    Fnd = False
    For P2 = 2 To Len(A0)
        C = Asc(Mid(A0, P2, 1))
        If AscIsLCase(C) Then Fnd = True: Exit For
    Next
'---
    If Not Fnd Then Exit Function
Dim P3%
Fnd = False
    For P3 = P2 + 1 To Len(A0)
        C = Asc(Mid(A0, P3, 1))
        If AscIsUCase(C) Or AscIsDigit(C) Then Fnd = True: Exit For
    Next
'--
If Fnd Then
    MthPfx = Left(A0, P3 - 1)
    Exit Function
End If
MthPfx = MthNm
End Function

Function FxWb(A) As Workbook
Set FxWb = Xls.Workbooks.Open(A)
End Function

Function FxaNm_Fxa$(A)
FxaNm_Fxa = CurPjPth & A & ".xlam"
End Function

Function HasPfx(S, Pfx) As Boolean
HasPfx = Left(S, Len(Pfx)) = Pfx
End Function

Function HasPfxS(S, Pfx) As Boolean
HasPfxS = Left(S, Len(Pfx) + 1) = Pfx & " "
End Function


Function HasSubStr(A, SubStr$) As Boolean
HasSubStr = InStr(A, SubStr) > 0
End Function



Function IsDigit(A) As Boolean
IsDigit = "0" <= A And A <= "9"
End Function

Function IsEmp(V) As Boolean
IsEmp = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsEmpty(V) Then Exit Function
If IsStr(V) Then
   If V = "" Then Exit Function
End If
If IsArray(V) Then
   If Sz(V) = 0 Then Exit Function
End If
IsEmp = False
End Function

Function IsFun(A As Mth) As Boolean
If Not IsStdMd(A.Md) Then Exit Function
IsFun = True
End Function

Function IsLetter(A) As Boolean
Dim C1$: C1 = UCase(A)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsMdNm(A) As Boolean
Select Case Left(A, 2)
Case "M_", "S_", "F_", "G_"
    IsMdNm = True
End Select
End Function

Function IsMthTy(A$) As Boolean
Select Case A
Case "Function", "Property Let", "Property Set", "Sub", "Function": IsMthTy = True
End Select
End Function

Function IsNm(A) As Boolean
If Not IsLetter(FstChr(A)) Then Exit Function
Dim L%: L = Len(A)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(A, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsNmChr(A$) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Function
If A = "_" Then Exit Function
If IsDigit(A) Then Exit Function
IsNmChr = False
End Function

Function IsNothing(A) As Boolean
IsNothing = TypeName(A) = "Nothing"
End Function

Function IsPfx(A, Pfx) As Boolean
IsPfx = Left(A, Len(Pfx)) = Pfx
End Function

Function IsPrim(A) As Boolean
Select Case VarType(A)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   IsPrim = True
End Select
End Function

Function IsPun(C) As Boolean
If IsLetter(C) Then Exit Function
If IsDigit(C) Then Exit Function
If C = "_" Then Exit Function
IsPun = True
End Function

Function IsStr(A) As Boolean
IsStr = VarType(A) = vbString
End Function

Function IsSy(A) As Boolean
IsSy = VarType(A) = vbArray + vbString
End Function
Function ItrAy(A)
ItrAy = ItrAyInto(A, Array())
End Function

Function ItrAyInto(A, OIntoAy)
Dim O: O = OIntoAy: Erase O
Dim I
For Each I In A
    Push O, I
Next
ItrAyInto = O
End Function

Function ItrNy(A)
Dim X, O$()
For Each X In A
    Push O, CallByName(X, "Name", VbGet)
Next
ItrNy = O
End Function
Function ItrNyWhPatnExl(A, Optional Patn$, Optional Exl$) As String()
ItrNyWhPatnExl = AyWhPatnExl(ItrNy(A), Patn, Exl)
End Function

Function JnCrLf$(A)
JnCrLf = Join(A, vbCrLf)
End Function

Function LasChr$(A)
LasChr = Right(A, 1)
End Function

Function IsCdLin(A) As Boolean
Dim L$: L = Trim(A)
If A = "" Then Exit Function
If Left(A, 1) = "'" Then Exit Function
IsCdLin = True
End Function
Function HasPfxAy(A, PfxAy) As Boolean
Dim P
For Each P In PfxAy
    If HasPfx(A, P) Then HasPfxAy = True: Exit Function
Next
End Function
Sub Z_LinMthKd()
Dim A$
Ept = "Property": A = "Private Property Get": GoSub Tst
Ept = "Property": A = "Property Get":         GoSub Tst
Ept = "Property": A = " Property Get":        GoSub Tst
Ept = "Property": A = "Friend Property Get":  GoSub Tst
Ept = "Property": A = "Friend  Property Get": GoSub Tst
Ept = "":         A = "FriendProperty Get":   GoSub Tst
Exit Sub
Tst:
    Act = LinMthKd(A)
    C
    Return
End Sub
Function LinMthKd$(A)
LinMthKd = TakMthKd(RmvMdy(A))
End Function

Function IsMthLin(A, Optional B As WhMth) As Boolean
If ObjPtr(B) = 0 Then
    IsMthLin = LinMthKd(A) <> ""
    Exit Function
End If
Dim M$, K$, N$, L$
L = A
M = ShfShtMdy(L): If Not AySel(B.InShtMdy, M) Then Exit Function
K = ShfMthKd(L): If K = "" Then Exit Function
If Not AySel(B.InShtKd, MthShtKd(K)) Then Exit Function
N = TakNm(L): If N = "" Then Stop
IsMthLin = IsNmSel(N, B.Nm)
End Function
Function IsNmSel(Nm$, B As WhNm) As Boolean
If Nm = "" Then Exit Function
If IsNothing(B) Then IsNmSel = True: Exit Function

End Function
Function LinIsTstSub(L$) As Boolean
LinIsTstSub = True
If IsPfx(L, "Sub Tst()") Then Exit Function
If IsPfx(L, "Sub Tst()") Then Exit Function
If IsPfx(L, "Friend Sub Tst()") Then Exit Function
If IsPfx(L, "Sub Z__Tst()") Then Exit Function
If IsPfx(L, "Sub Z__Tst()") Then Exit Function
If IsPfx(L, "Friend Sub Z__Tst()") Then Exit Function
LinIsTstSub = False
End Function

Function LinMayLCC(A, MthNm$, Lno%) As MayLCC
Dim M$: M = LinMthNm(A)
If M <> MthNm Then Set LinMayLCC = NonLCC: Exit Function
Dim C1%, C2%
C1 = InStr(A, MthNm)
C2 = C1 + Len(MthNm)
Set LinMayLCC = SomLCC(LCC(Lno, C1, C2))
End Function

Function TakMdy$(A)
TakMdy = TakPfxAyS(A, MdyAy)
End Function

Function ShfMdy$(OLin$)
Dim O$
O = TakMdy(OLin): If O = "" Then Exit Function
ShfMdy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function LinMthTy$(A)
LinMthTy = TakPfxAyS(RmvMdy(A), MthTyAy)
End Function

Function LinMthShtTy$(A)
LinMthShtTy = MthShtTy(LinMthTy(A))
End Function

Function ShfNm$(OLin$)
Dim O$: O = TakNm(OLin): If O = "" Then Exit Function
ShfNm = O
OLin = RmvPfx(OLin, O)
End Function

Function TakNm$(A)
Dim J%
If Not IsLetter(Left(A, 1)) Then Exit Function
For J = 2 To Len(A)
    If Not IsNmChr(Mid(A, J, 1)) Then
        TakNm = Left(A, J - 1)
        Exit Function
    End If
Next
TakNm = A
End Function


Function PfxAyFstS$(PfxAy$(), Lin) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
Dim P
For Each P In PfxAy
    If HasPfx(Lin, P) Then If Mid(Lin, Len(P) + 1, 1) = " " Then PfxAyFstS = P: Exit Function
Next
End Function

Function PfxAyFst$(PfxAy$(), Lin) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
Dim P
For Each P In PfxAy
    If HasPfx(Lin, P) Then PfxAyFst = P: Exit Function
Next
End Function

Function TakPfx$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx else return ""
If HasPfx(Lin, Pfx) Then TakPfx = Pfx
End Function

Function TakPfxS$(Lin, Pfx$) ' Return [Pfx] if [Lin] has such pfx+" " else return ""
If HasPfx(Lin, Pfx) Then If Mid(Lin, Len(Pfx) + 1, 1) = " " Then TakPfxS = Pfx
End Function

Function TakPfxAyS$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P and a space
TakPfxAyS = PfxAyFstS$(PfxAy, Lin)
End Function

Function TakPfxAy$(Lin, PfxAy$()) ' Return Fst ele-P of [PfxAy] if [Lin] has pfx ele-P
TakPfxAy = PfxAyFst$(PfxAy, Lin)
End Function

Function MthKd$(MthTy$)
Select Case MthTy
Case "Function": MthKd = "Fun"
Case "Sub": MthKd = "Sub"
Case "Property Get", "Property Get", "Property Let": MthKd = "Prp"
End Select
End Function
Function MthShtTyKd$(MthShtTy$)
Select Case MthShtTy
Case "Fun", "Sub": MthShtTyKd = MthShtTy
Case "Get", "Let", "Set": MthShtTyKd = "Prp"
End Select
End Function

Function RmvMdy$(A)
RmvMdy = LTrim(RmvPfxAyS(A, MdyAy))
End Function

Function LinRmvT1$(A)
Dim O$: O = A
ShfTerm O
LinRmvT1 = O
End Function
Function TakBet$(A, S1$, S2$)
Dim P%, L%, P1%, P2%
P1 = InStr(A, S1): If P1 = 0 Then Exit Function
P = P1 + Len(S1)
P2 = InStr(P, A, S2): If P2 = 0 Then Exit Function
L = P2 - P1 - 1
TakBet = Mid(A, P, L)
End Function
Function ApIntAy(ParamArray Ap()) As Integer()
Dim Av(): Av = Ap
ApIntAy = AyIntAy(Av)
End Function


Function ShfShtMdy$(OLin$)
ShfShtMdy = ShtMdy(ShfMdy(OLin))
End Function
Function ShfMthKd$(OLin$)
Dim T$: T = TakMthKd(OLin)
If T = "" Then Exit Function
ShfMthKd = T
OLin = LTrim(RmvPfx(OLin, T))
End Function
Function ShfMthShtKd$(OLin$)
Dim O$
O = ShfMthKd(OLin$): If O = "" Then Exit Function
ShfMthShtKd = MthShtKd(O)
End Function

Function ShfMthTy$(OLin$)
Dim O$: O = TakMthTy(OLin)
If O = "" Then Exit Function
ShfMthTy = O
OLin = LTrim(RmvPfx(OLin, O))
End Function

Function ShfMthShtTy$(OLin$)
ShfMthShtTy = MthShtTy(ShfMthTy(OLin))
End Function

Private Sub Z_ShfVal()
Dim A$, SSNm$
A = "": SSNm = "": Ept = Array(Array(), ApSy()): GoSub Tst
Exit Sub
Tst:
    Act = ShfVal(A, SSNm)
    C
    Return
End Sub

Function LinTermAy(A) As String()
Dim L$, T$
L = A
X:
If L = "" Then Exit Function
T = ShfTerm(L): If T = "" Then Stop
Push LinTermAy, T
GoTo X
End Function

Function ShfStarTerm(OItm$(), OLbl$()) As Variant()
Dim NStar%, I
For Each I In OLbl
    If FstChr(I) <> "*" Then
        If NStar > 0 Then
            OItm = AyMid(OItm, NStar)
            OLbl = AyMid(OLbl, NStar)
            Exit Function
        End If
    End If
    Push ShfStarTerm, OItm(NStar)
    NStar = NStar + 1
Next
End Function

Function ShfLblVal(OItm$(), OLblQ) As Variant()

End Function
Function ShfVal(Lin$, SSLbl$) As Variant()
Dim Lbl$(), Itm$(), Lbli
Lbl = SslSy(SSLbl)
Itm = LinTermAy(Lin)
ShfVal = ShfStarTerm(Itm, Lbl)
For Each Lbli In AyNz(Lbl)
    If Sz(Itm) = 0 Then Exit For
    PushI ShfVal, ShfLblVal(Itm, Lbli)
Next
End Function

Sub Z_ShfTerm()
Dim A$, Ept1$
A = " AA BB "
Ept = "AA"
Ept1 = "BB "
GoSub Tst
Exit Sub
Tst:
    Act = ShfTerm(A)
    C
    Ass A = Ept1
    Return
End Sub
Function TakT1$(A)
If FstChr(A) <> "[" Then TakT1 = TakBef(A, " "): Exit Function
Dim P%
P = InStr(A, "]")
If P = 0 Then Stop
TakT1 = Mid(A, 2, P - 2)
End Function

Function ShfX$(OLin$, X$)
If LinT1(OLin) = X Then
    ShfX = X
    OLin = LinRmvT1(OLin)
    Exit Function
End If
End Function

Function ShfMthSfx$(OLin$)
Const C$ = "#!@#$%^&"
Dim F$, P%
F = FstChr(OLin)
P = InStr(C, F)
If P > 0 Then
    ShfMthSfx = Mid(C, P, 1)
    OLin = Mid(OLin, 2)
    Exit Function
End If
End Function
Function LinT1$(A)
LinT1 = TakT1(LTrim(A))
End Function

Function LinesAyFmt(A$()) As String()
Dim LyAy()
    LyAy = AyMap(A, "SplitCrLf")
Dim W%()
    W = AyMapInto(LyAy, "AyWdt", EmpIntAy)
Dim NRowAy%()
    NRowAy = AyMapInto(LyAy, "Sz", EmpIntAy)
Dim NRow%
    NRow = AyMax(NRowAy)
Dim O$()
    Dim J%, Hdr$
    Hdr = WdtAy_HdrLin(W)
    Push O, Hdr
    For J = 0 To NRow - 1
        Push O, LyAy_Lin(LyAy, W, J)
    Next
    Push O, Hdr
LinesAyFmt = O
End Function

Function LinesAyWdt%(A)
If Sz(A) = 0 Then Exit Function
Dim O%, J&, M%, L
For Each L In A
   O = Max(O, LinesWdt(L))
Next
LinesAyWdt = O
End Function

Function LinesBoxLy(A) As String()
LinesBoxLy = LyBoxLy(SplitCrLf(A))
End Function
Function LinCnt&(Lines)
LinCnt = SubStrCnt(Lines, vbCrLf) + 1
End Function

Function LinesSqV(Lines$) As Variant
LinesSqV = AySqV(SplitCrLf(Lines))
End Function

Function LinesTrimEnd$(A$)
LinesTrimEnd = Join(LyTrimEnd(SplitCrLf(A)), vbCrLf)
End Function

Function LinesUnderLin$(Lines)
LinesUnderLin = StrDup("-", LinesWdt(Lines))
End Function

Function LinesVbl$(A)
LinesVbl = Replace(A, vbCrLf, "|")
End Function

Function LinesWdt%(A)
LinesWdt = AyWdt(SplitCrLf(A))
End Function

Function LoQt(A As ListObject) As QueryTable
On Error Resume Next
Set LoQt = A.QueryTable
End Function

Function LyAy_Lin$(A(), WdtAy%(), Ix%)
Dim J%, W%, I$, Ly$(), Dr$()
For J = 0 To UB(A)
    Ly = A(J)
    W% = WdtAy(J)
    If UB(Ly) >= Ix Then
        I = Ly(Ix)
    Else
        I = ""
    End If
    Push Dr, AlignL(I, W)
Next
LyAy_Lin = "| " + Join(Dr, " | ") + " |"
End Function

Function LyBoxLy(A) As String()
If Sz(A) = 0 Then Exit Function
Dim W%: W = AyWdt(A)
Dim H$: H = "|" & StrDup("-", W + 2) & "|"
Dim O$()
Push O, H
Dim I
For Each I In A
    Push O, "| " & AlignL(I, W) + " |"
Next
Push O, H
LyBoxLy = O
End Function

Function LyTrimEnd(Ly) As String()
If Sz(Ly) = 0 Then Exit Function
Dim L$
Dim J&
For J = UB(Ly) To 0 Step -1
    L = Trim(Ly(J))
    If Trim(Ly(J)) <> "" Then
        Dim O$()
        O = Ly
        ReDim Preserve O(J)
        LyTrimEnd = O
        Exit Function
    End If
Next
End Function

Function Max(A, B)
If A > B Then
    Max = A
Else
    Max = B
End If
End Function

Function MaxCol&()
Static C&, Y As Boolean
If Not Y Then
    Y = True
    C = IIf(Application.Version = "16.0", 16384, 255)
End If
MaxCol = C
End Function

Function MaxRow&()
Static R&, Y As Boolean
If Not Y Then
    Y = True
    R = IIf(Application.Version = "16.0", 1048576, 65535)
End If
MaxRow = R
End Function

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


Function MdSrtDic(A As CodeModule) As Dictionary
Set MdSrtDic = DicAddKeyPfx(SrcSrtDic(MdSrc(A)), MdNm(A) & ".")
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
Function IsMdAllRemarked(A As CodeModule) As Boolean
Dim J%, L$
For J = 1 To A.CountOfLines
    If Left(A.Lines(J, 1), 1) <> "'" Then Exit Function
Next
IsMdAllRemarked = True
End Function

Function IsClsMd(A As CodeModule) As Boolean
IsClsMd = A.Parent.Type = vbext_ct_ClassModule
End Function

Function IsStdMd(A As CodeModule) As Boolean
IsStdMd = A.Parent.Type = vbext_ct_StdModule
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

Function MdMthAy(A As CodeModule, Optional B As WhMth) As Mth()
Dim N
For Each N In AyNz(MdMthNy(A, B))
    PushObj MdMthAy, Mth(A, N)
Next
End Function

Function MdMthLno(A As CodeModule, MthNm) As Integer()
MdMthLno = AyAdd1(SrcMthNmIx(MdSrc(A), MthNm))
End Function

Function MdMthSq(A As CodeModule) As Variant()
MdMthSq = MthKy_Sq(MdMthKy(A, True))
End Function

Function PjMthSq(A As VBProject) As Variant()
PjMthSq = MthKy_Sq(PjMthKy(A, True))
End Function

Function CurMdMthNy(Optional A As WhMth) As String()
CurMdMthNy = MdMthNy(CurMd, A)
End Function

Function MdMthNy(A As CodeModule, Optional B As WhMth) As String()
MdMthNy = AyAddPfx(SrcMthNy(MdBdyLy(A), B), MdShtTy(A) & "." & MdNm(A) & ".")
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
Dim X As Dictionary, Y As Dictionary
Set X = SrcDic(MdSrtedLy(A))
Set Y = SrcDic(MdSrc(A))
MdSrtRpt = DicCmp(X, Y, "BefSrt", "AftSrt")
End Function

Function CurMdSrtRptFmt() As String()
CurMdSrtRptFmt = MdSrtRptFmt(CurMd)
End Function
Function MdSrtRptFmt(A As CodeModule) As String()
MdSrtRptFmt = DCRsltFmt(MdSrtRpt(A))
End Function
Sub Mov()
CurMthMov "IdeSrt"
End Sub

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

Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
Min = AyMin(Av)
End Function

Function MthDNm$(A As Mth)
MthDNm = MdDNm(A.Md) & "." & A.Nm
End Function

Function MthDNmLines$(A)
MthDNmLines = MthLines(DMth(A))
End Function

Function MthFul$(MthNm)
MthFul = VbeMthMdDNm(CurVbe, MthNm)
End Function

Function MthFuly(MthNm) As String()
MthFuly = VbeMthMdDNy(CurVbe, MthNm)
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
Function MthNmMdDNy(A) As String()
MthNmMdDNy = CurVbeMthMdDNy(A)
End Function
Function CurVbeMthMdDNy(MthNm) As String()
CurVbeMthMdDNy = VbeMthMdDNy(CurVbe, MthNm)
End Function
Function MthNmMd(A) As CodeModule '
Dim O As CodeModule
Set O = CurMd
If MdHasMth(O, A) Then Set MthNmMd = O: Exit Function
Dim N$
N = MthFul(A)
If N = "" Then
    Debug.Print FmtQQ("Mth[?] not found in any Pj")
    Exit Function
End If
Set MthNmMd = Md(N)
End Function
Function DMth(MthDNm) As Mth
Dim Ay$(): Ay = Split(MthDNm, ".")
Dim Nm$, M As CodeModule
Select Case Sz(Ay)
Case 1: Nm = Ay(0): Set M = MthNmMd(Ay(0))
Case 2: Nm = Ay(1): Set M = Md(Ay(0))
Case 3: Nm = Ay(2): Set M = Md(Ay(0) & "." & Ay(1))
Case Else: Stop
End Select
Set DMth = Mth(M, Nm)
End Function

Function MthDNm_Nm$(A)
Dim Ay$(): Ay = Split(A, ".")
Dim Nm$
Select Case Sz(Ay)
Case 1: Nm = Ay(0)
Case 2: Nm = Ay(1)
Case 3: Nm = Ay(2)
Case Else: Stop
End Select
MthDNm_Nm = Nm
End Function

Function MthFNm$(A As Mth)
MthFNm = A.Nm & ":" & MdDNm(A.Md)
End Function

Function MthFNm_Mth(A) As Mth
Set MthFNm_Mth = DMth(MthFNm_MthDNm(A))
End Function

Function MthFNm_MthDNm$(A)
With Brk(A, ":")
    MthFNm_MthDNm = .S2 & "." & .S1
End With
End Function

Function MthFNm_Nm$(A$)
MthFNm_Nm = Brk(A, ":").S1
End Function

Function MthLno(A As Mth) As Integer()
MthLno = MdMthLno(A.Md, A.Nm)
End Function

Function MthLnoAy(A As Mth) As Integer()
MthLnoAy = AyAdd1(SrcMthNmIx(MdSrc(A.Md), A.Nm))
End Function

Function MthRmkFC(A As Mth) As FmCnt()
MthRmkFC = SrcMthRmkFC(MdSrc(A.Md), A.Nm)
End Function

Function IsMthExist(A As Mth) As Boolean
IsMthExist = MdHasMth(A.Md, A.Nm)
End Function

Function IsPubMth(A As Mth) As Boolean
Dim L$: L = MthLin(A): If L = "" Then Stop
Dim Mdy$: Mdy = TakMdy(L)
If Mdy = "" Or Mdy = "Public" Then IsPubMth = True
End Function

Function MthKy_Sq(A$()) As Variant()
Dim O(), J%
ReDim O(1 To Sz(A) + 1, 1 To 6)
SqSetRow O, 1, MthKeyDrFny
For J = 0 To UB(A)
    SqSetRow O, J + 2, Split(A(J), ":")
Next
MthKy_Sq = O
End Function

Function MthMayLCC(A As Mth) As MayLCC
Dim L%, C As MayLCC
Dim M As CodeModule
Set M = A.Md
For L = M.CountOfDeclarationLines + 1 To M.CountOfLines
    Set C = LinMayLCC(M.Lines(L, 1), A.Nm, L)
    If C.Som Then
        Set MthMayLCC = SomLCC(C.LCC)
        Exit Function
    End If
Next
End Function
Function MthANy_MthNmItr(A) As Collection
Dim O As New Collection, J&
For J = 0 To UB(A)
    ItrPushNoDup O, A(J)
Next
Set MthANy_MthNmItr = O
End Function
Function ItrHas(A As Collection, M) As Boolean
Dim I
For Each I In A
    If I = M Then ItrHas = True: Exit Function
Next
End Function
Sub ItrPushNoDup(A As Collection, M)
If ItrHas(A, M) Then Exit Sub
A.Add M
End Sub
Function MthLin$(A As Mth)
MthLin = SrcMthLin(MdBdyLy(A.Md), A.Nm)
End Function
Function MthANm_MthNm$(A)
MthANm_MthNm = TakBefOrAll(A, ":")
End Function
Function MthBNm_MthNm$(A)
MthBNm_MthNm = MthANm_MthNm(MthBNm_MthANm(A))
End Function
Function MthBNm_MdNm$(A)
MthBNm_MdNm = TakBefMust(A, ".")
End Function

Function MthBNm_MthANm$(A)
MthBNm_MthANm = TakAftMust(A, ".")
End Function

Function CvFmCnt(A) As FmCnt
Set CvFmCnt = A
End Function

Function FmCntAyLinCnt%(A() As FmCnt)
Dim I, C%, O%
For Each I In A
    C = CvFmCnt(I).Cnt
    If C > 0 Then O = O + C
Next
FmCntAyLinCnt = O
End Function

Function MthLinCnt%(A As Mth)
MthLinCnt = FmCntAyLinCnt(MthFC(A))
End Function
Sub Z_LinMthNm()
GoTo ZZ
Dim A$
A = "Function LinMthNm$(A)": Ept = "LinMthNm.Fun.": GoSub Tst
Exit Sub
Tst:
    Act = LinMthNm(A)
    C
    Return
ZZ:
    Dim O$(), L, P, M
    For Each P In VbePjAy(CurVbe)
        For Each M In PjMdAy(CvPj(P))
            For Each L In MdBdyLy(CvMd(M))
                PushNonBlankStr O, LinMthNm(L, WhMth("Prv"))
            Next
        Next
    Next
    Brw O
End Sub
Sub Z_FbMthNy()
GoSub X_BrwAll
Exit Sub
X_BrwAll:
    Dim O$(), Fb
    For Each Fb In AppFbAy
        PushAy O, FbMthNy(Fb)
    Next
    Brw O
    Return
X_BrwOne:
    Brw FbMthNy(AppFbAy()(0))
    Return
End Sub
Function FbMthNy(A) As String()
FbMthNy = VbeMthNy(FbAcs(A).Vbe)
End Function

Function LinMthNm$(A, Optional B As WhMth)
Dim L$, M$, T$, N$, Sel As Boolean
Sel = Not IsNothing(B)
L = A
Dim NotSel As Boolean
M = ShfShtMdy(L):   GoSub X1: If NotSel Then Exit Function
T = ShfMthTy(L):    GoSub X2: If NotSel Then Exit Function
N = TakNm(L):       GoSub X3: If NotSel Then Exit Function
LinMthNm = N & "." & MthShtTy(T) & "." & M
Exit Function
X1: If Sel Then NotSel = Not AySel(B.InShtMdy, M)
    Return
X2: If T = "" Then NotSel = True: Return
    If Sel Then NotSel = Not AySel(B.InShtKd, MthShtTyKd(T))
    Return
X3:
    If N = "" Then Stop
    If Sel Then NotSel = Not IsNmSel(N, B.Nm)
    Return
End Function

Function MthNmSrtKey$(A)
If A = "*Dcl" Then MthNmSrtKey = "0.*Dcl": Exit Function
Dim B$(): B = SplitDot(A)
Dim N$
Dim M$
Dim T$
    If B(1) = "" Then Stop
    If B(0) = "" Then Stop
    N = B(0)
    T = B(1)
    M = B(2)
Dim P% 'Priority
    Select Case True
    Case IsPfx(N, "Init"): P = 1
    Case N = "Z__Tst":     P = 9
    Case N = "ZZ__Tst":    P = 9
    Case IsPfx(N, "Z_"):   P = 9
    Case IsPfx(N, "ZZ_"):  P = 8
    Case IsPfx(N, "Z"):    P = 7
    Case Else:             P = 2
    End Select
MthNmSrtKey = P & "." & A
End Function

Function MthMdNm$(A As Mth)
MthMdNm = MdNm(A.Md)
End Function
Function MthMdDNm$(A As Mth)
MthMdDNm = MdDNm(A.Md)
End Function

Function MthNm$(A As Mth)
MthNm = A.Nm
End Function

Function MthPjNm$(A As Mth)
MthPjNm = MdPjNm(A.Md)
End Function

Function MthShtKd$(MthKd$)
Dim O$
Select Case MthKd
Case "Sub": O = MthKd
Case "Function": O = "Fun"
Case "Property": O = "Prp"
End Select
MthShtKd = O
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

Function NewA1(Optional WsNm$ = "Sheet1") As Range
Set NewA1 = NewWs(WsNm).Range("A1")
End Function

Function NewXls() As Excel.Application
Set NewXls = New Excel.Application
End Function

Function NewWb(Optional Ws$ = "Sheet1") As Workbook
Dim O As Workbook, W As Worksheet
Set O = NewXls.Workbooks.Add
Set W = O.Sheets(1)
If W.Name <> Ws Then W.Name = Ws
Set NewWb = O
End Function

Function NewWs(Optional WsNm$) As Worksheet
Set NewWs = WsSetNm(NewWb.Sheets(1), WsNm)
End Function

Function OyPrpAy(Oy, PrpNm) As Variant()
OyPrpAy = OyPrpAyInto(Oy, PrpNm, EmpAy)
End Function

Function OyPrpAyInto(Oy, PrpNm, OIntoAy)
Dim O: O = OIntoAy: Erase O
If Sz(Oy) > 0 Then
    Dim I
    For Each I In Oy
        Push O, ObjPrp(I, PrpNm)
    Next
End If
OyPrpAyInto = O
End Function

Function ObjPrp(Obj, PrpNm)
On Error Resume Next
Asg CallByName(Obj, PrpNm, VbGet), ObjPrp
End Function


Function OyNy(Oy) As String()
Dim O$(): If Sz(Oy) = 0 Then Exit Function
Dim I
For Each I In Oy
    Push O, CallByName(I, "Name", VbGet)
Next
OyNy = O
End Function

Function OyToStrSy(A) As String()
If Sz(A) = 0 Then Exit Function
Dim O$()
ReDim O(UB(A))
Dim J&
For J = 0 To UB(A)
    O(J) = A(J).ToStr
Next
OyToStrSy = O
End Function
Private Function Z_ShfXXX()
Dim O$: O = "AA{|}BB "
Ass ShfXXX(O, "{|}") = "AA"
Ass O = "BB "
End Function
Function ShfXXX$(O$, XXX$)
Dim P%: P = InStr(O, XXX)
If P = 0 Then Exit Function
ShfXXX = Left(O, P - 1)
O = Mid(O, P + Len(XXX))
End Function
Function ShfDTerm$(O$)
ShfDTerm = ShfXXX(O, ".")
End Function

Function PjNy() As String()
PjNy = ItrNy(CurVbe.VBProjects)
End Function
Function IsPjNm(A) As Boolean
IsPjNm = AyHas(PjNy, A)
End Function
Function DicAddKeyPfx(A As Dictionary, Pfx) As Dictionary
Dim O As New Dictionary, K
For Each K In A.Keys
    O.Add Pfx & K, A(K)
Next
Set DicAddKeyPfx = O
End Function
Function Pj(PjNm) As VBProject
Set Pj = CurVbe.VBProjects(PjNm)
End Function
Function SelOy(A, PrpSsl$) As Variant()

End Function

Function OyWhPrpIn(A, P, InAy)
Dim X, O
If Sz(A) = 0 Or Sz(InAy) Then OyWhPrpIn = A: Exit Function
O = A
Erase O
For Each X In A
    If AyHas(InAy, ObjPrp(X, P)) Then PushObj O, X
Next
OyWhPrpIn = O
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

Function PjClsAy(A As VBProject, Optional B As WhNm) As CodeModule()
'123
'PjClsAy = PjMdAy(A,IdeNew.WhNm WhMd(, SelClsMd)
End Function

Function SelStdMd() As WhMd
SelStdMd = WhMd("Std")
End Function
Function SelClsMd() As WhMd
SelClsMd = WhMd("Cls")
End Function
Function ClsCmp() As vbext_ComponentType()
ClsCmp = CvWhCmpTy("Cls")
End Function
Function StdCmp() As vbext_ComponentType()
StdCmp = CvWhCmpTy("Std")
End Function
Function WhEmpNm() As WhNm
End Function
Function PjStdAy(A As VBProject, Optional B As WhNm) As CodeModule()
'123
'PjStdAy = PjMdAyWh(A, SelStdMd)
End Function

Function OyWhNm(A, B As WhNm)
Dim X
For Each X In AyNz(A)
    If IsNmSel(X.Name, B) Then PushObj OyWhNm, X
Next
End Function

Function PjMdNy(A As VBProject, Optional B As WhMd) As String()
PjMdNy = PjCmpNy(A, B)
End Function
Function PjCmpNy(A As VBProject, Optional B As WhMd) As String()
PjCmpNy = ItrNy(PjCmpAy(A, B))
End Function

Function PjClsNy(A As VBProject, Optional B As WhNm) As String()
PjClsNy = PjCmpNy(A, WhMd("Cls", B))
End Function

Function PjCmp(A As VBProject, Nm) As VBComponent
Set PjCmp = A.VBComponents(Nm)
End Function

Function PjDic(A As VBProject) As Dictionary
Dim I
Dim O As New Dictionary
For Each I In PjMdAy(A)
    Set O = DicAdd(O, MdDic(CvMd(I)))
Next
Set PjDic = O
End Function

Function PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.Filename
End Function

Function PjFstMd(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If Cmp.Type = vbext_ct_StdModule Then
        Set PjFstMd = Cmp.CodeModule
        Exit Function
    End If
Next
End Function
Function PjFstMbr(A As VBProject) As CodeModule
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    Set PjFstMbr = Cmp.CodeModule
    Exit Function
Next
End Function

Function PjFunBdyDic(A As VBProject) As Dictionary
Stop '
End Function

Function CurPjFunPfxAy() As String()
CurPjFunPfxAy = PjFunPfxAy(CurPj)
End Function

Function PjFunPfxAy(A As VBProject) As String()
Dim Ay() As CodeModule: Ay = PjMdAy(A)
Dim Ay1(): Ay1 = AyMap(Ay, "MdFunPfx")
PjFunPfxAy = AyFlat(Ay1)
End Function

Sub LocStr_Go(A)
LocGo LocStr_Loc(A)
End Sub
Function LocStr_Loc(A) As Loc

End Function
Sub LocGo(A As Loc)

End Sub

Function ItrMap(A, MapFunNm$)
Dim I, O As New Collection
For Each I In A
    O.Add Run(MapFunNm, I)
Next
Set ItrMap = O
End Function

Function ItrMapSy(A, MapFunNm$) As String()
ItrMapSy = ItrSy(ItrMap(A, MapFunNm))
End Function

Function PjRfNy(A As VBProject) As String()
PjRfNy = ItrNy(A.References)
End Function
Function PjHasRfNm(A As VBProject, RfNm$) As Boolean
Dim I, R As Reference
For Each I In A.References
    Set R = I
    If R.Name = RfNm Then PjHasRfNm = True: Exit Function
Next
End Function
Function PjHasRfFfn(A As VBProject, RfFfn) As Boolean
Dim I, R As Reference
For Each I In A.References
    Set R = I
    If R.FullPath = RfFfn Then PjHasRfFfn = True: Exit Function
Next
End Function
Function ItrWhNm(A, B As WhNm)
ItrWhNm = ItrWhNmInto(A, B, EmpAy)
End Function

Function ItrWhNmInto(A, B As WhNm, OInto)
Erase OInto
Dim X
For Each X In A
    If IsNmSel(X.Name, B) Then PushObj OInto, X
Next
ItrWhNmInto = OInto
End Function

Function MdTy(A As CodeModule) As vbext_ComponentType
MdTy = A.Parent.Type
End Function

Function CvCmpTyAy(CmpTyAy0$) As vbext_ComponentType()
Dim X, O() As vbext_ComponentType
For Each X In SslSy(CmpTyAy0)
    Push O, CmpShtToTy(X)
Next
CvCmpTyAy = O
End Function

Function CurPjMdAy() As CodeModule()
CurPjMdAy = PjMdAy(CurPj)
End Function

Function PjMdAy(A As VBProject, Optional B As WhMd) As CodeModule()
If IsNothing(B) Then
    PjMdAy = ItrPrpAyInto(A.VBComponents, "CodeModule", PjMdAy)
    Exit Function
End If
Dim C
For Each C In AyNz(ItrWhNm(A.VBComponents, B.Nm))
    With CvCmp(C)
        If AySel(B.InCmpTy, .Type) Then
            PushObj PjMdAy, .CodeModule
        End If
    End With
Next
End Function

Function ItrWhInNyInto(A, InNy$(), OInto)
Dim O, X
O = OInto
Erase O
For Each X In A
    If AyHas(InNy, X.Name) Then PushObj O, X
Next
ItrWhInNyInto = O
End Function

Function CvWhCmpTy(WhCmpTy$) As vbext_ComponentType()
Dim O() As vbext_ComponentType, I
For Each I In AyNz(SslSy(WhCmpTy))
    Push O, CmpShtToTy(I)
Next
CvWhCmpTy = O
End Function

Function PjCmpAy(A As VBProject, Optional B As WhMd) As VBComponent()
If IsNothing(B) Then
    PjCmpAy = ItrAyInto(A.VBComponents, PjCmpAy)
    Exit Function
End If
Dim Cmp
For Each Cmp In AyNz(ItrWhNm(A.VBComponents, B.Nm))
    If AySel(B.InCmpTy, CvCmp(Cmp).Type) Then PushObj PjCmpAy, Cmp
Next
End Function

Private Sub ZZ_PjHasMd()
Ass PjHasMd(CurPj, "Drs") = False
Ass PjHasMd(CurPj, "A__Tool") = True
End Sub

Function PjNm$(A As VBProject)
PjNm = A.Name
End Function
Function PjMdOpt(A As VBProject, Nm) As CodeModule
If Not PjHasMd(A, Nm) Then Exit Function
Set PjMdOpt = PjMd(A, Nm)
End Function
Function PjMd(A As VBProject, Nm) As CodeModule
Set PjMd = PjCmp(A, Nm).CodeModule
End Function

Function PjStdNy(A As VBProject, Optional B As WhNm) As String()
PjStdNy = PjCmpNy(A, WhMd("Std", B))
End Function

Function PjMdNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_StdModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
PjMdNy_With_TstSub = O
End Function

Function PjMdSrtRptDic(A As VBProject) As Dictionary 'Return a dic of [MdNm,SrtCmpFmt]
'SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
Dim I, O As New Dictionary, Md As CodeModule
    For Each I In AyNz(PjMdAy(A))
        Set Md = I
        O.Add MdNm(Md), MdSrtRptFmt(CvMd(Md))
    Next
Set PjMdSrtRptDic = O
End Function

Function PjStdClsNy(A As VBProject, Optional B As WhNm) As String()
PjStdClsNy = PjCmpNy(A, WhMd("Std Cls", B))
End Function


Function PjMthKy(A As VBProject, Optional IsWrap As Boolean) As String()
PjMthKy = AyMapPXSy(PjMdAy(A), "MdMthKy", IsWrap)
End Function

Function PjMthKySq(A As VBProject) As Variant()
PjMthKySq = MthKy_Sq(PjMthKy(A, True))
End Function

Function CurPjMthNy(Optional A As WhPjMth) As String()
CurPjMthNy = PjMthNy(CurPj, A)
End Function

Function PjMthNy(A As VBProject, Optional B As WhMdMth) As String()
Dim Md, N$, Ny$(), WMd As WhMd, WMth As WhMth
Set WMd = WhMdMth_Md(B)
Set WMth = WhMdMth_Mth(B)
N = A.Name & "."
For Each Md In AyNz(PjMdAy(A, WMd))
    Ny = MdMthNy(CvMd(Md), WMth)
    Ny = AyAddPfx(Ny, N)
    PushAyNoDup PjMthNy, Ny
Next
End Function

Function PjPth$(A As VBProject)
PjPth = FfnPth(A.Filename)
End Function

Function PjRfAy(A As VBProject) As Reference()
PjRfAy = ItrAyInto(A.References, EmpRfAy)
End Function

Function PjRfCfgFfn(A As VBProject)
PjRfCfgFfn = PjSrcPth(A) & "PjRf.Cfg"
End Function

Function PjRfLy(A As VBProject) As String()
Dim RfAy() As Reference
    RfAy = PjRfAy(A)
Dim O$()
Dim Ny$(): Ny = OyNy(RfAy)
Ny = AyAlignL(Ny)
Dim J%
For J = 0 To UB(Ny)
    Push O, Ny(J) & " " & RfFfn(RfAy(J))
Next
PjRfLy = O
End Function

Function PjSrcPth(A As VBProject)
Dim Ffn$: Ffn = PjFfn(A)
If Ffn = "" Then Exit Function
Dim Fn$: Fn = FfnFn(Ffn)
Dim P$: P = FfnPth(A.Filename)
If P = "" Then Exit Function
Dim O$:
O = P & "Src\": PthEns O
O = O & Fn & "\":                  PthEns O
PjSrcPth = O
End Function

Function PjSrtRptFmt(A As VBProject) As String()
Dim O$(), I
For Each I In AyNz(PjMdAy(A))
    PushAy O, MdSrtRptFmt(CvMd(I))
Next
PjSrtRptFmt = O
End Function

Function LyDicWb(A As Dictionary) As Workbook

End Function

Function PjSrtRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
Set PjSrtRptWb = LyDicWb(PjMdSrtRptDic(A))
Stop '
Dim O As Workbook: ' Set O = DicWb(A1.RptDic)
Dim Ws As Worksheet
Set Ws = WbAddWs(O, "Md Idx")
'Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
'LoCol_LnkWs Lo, "Md"
'If Vis Then WbVis O
'Set PjSrtRptWb = O
Stop '
End Function

Function Pj_ClsNy_With_TstSub(A As VBProject) As String()
Dim I As VBComponent
Dim O$()
For Each I In A.VBComponents
    If I.Type = vbext_ct_ClassModule Then
        If MdHasTstSub(I.CodeModule) Then
            Push O, I.Name
        End If
    End If
Next
Pj_ClsNy_With_TstSub = O
End Function

Function Pj_TstClass_Bdy$(A As VBProject)
Dim N1$() ' All Class Ny with 'Friend Sub Z__Tst' method
Dim N2$()
Dim A1$, A2$
Const Q1$ = "Sub ?()|Dim A As New ?: A.Z__Tst|End Sub"
Const Q2$ = "Sub ?()|#.?.Z__Tst|End Sub"
N1 = Pj_ClsNy_With_TstSub(A)
A1 = SeedExpand(Q1, N1)
N2 = PjMdNy_With_TstSub(A)
A2 = Replace(SeedExpand(Q2, N2), "#", A.Name)
Pj_TstClass_Bdy = A1 & vbCrLf & A2
End Function

Function Re(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
If Patn = "" Then Exit Function
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set Re = O
End Function

Function RfFfn$(A As Reference)
On Error Resume Next
RfFfn = A.FullPath
End Function

Function PjRfNm_RfFfn$(A As VBProject, RfNm$)
PjRfNm_RfFfn = PjPth(A) & RfNm & ".xlam"
End Function

Function RgLo(A As Range, Optional LoNm$) As ListObject
Dim O As ListObject
Set O = RgWs(A).ListObjects.Add(xlSrcRange, A, , xlYes)
If LoNm <> "" Then O.Name = LoNm
Set RgLo = O
End Function

Function RgRC(A As Range, R, C) As Range
Set RgRC = A.Cells(R, C)
End Function

Function RgRCRC(A As Range, R1, C1, R2, C2) As Range
Set RgRCRC = RgWs(A).Range(RgRC(A, R1, C1), RgRC(A, R2, C2))
End Function

Function RgWs(A As Range)
Set RgWs = A.Parent
End Function

Function RgWb(A As Range)
Set RgWb = WsWb(RgWs(A))
End Function

Function RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Function

Function RmvLasChr$(A)
RmvLasChr = Left(A, Len(A) - 1)
End Function

Function RmvLasNChr$(A, N%)
RmvLasNChr = Left(A, Len(A) - N)
End Function

Function RmvPfx$(A, Pfx)
If IsPfx(A, Pfx) Then
    RmvPfx = Mid(A, Len(Pfx) + 1)
Else
    RmvPfx = A
End If
End Function

Function RplDblSpc$(A)
Dim O$: O = Trim(A)
Dim J&
While HasSubStr(O, "  ")
    J = J + 1: If J > 10000 Then Stop
    O = Replace(O, "  ", " ")
Wend
RplDblSpc = O
End Function

Function RplPun$(A)
Dim O$(), J&, L&, C$
L = Len(A)
If L = 0 Then Exit Function
ReDim O(L - 1)
For J = 1 To L
    C = Mid(A, J, 1)
    If IsPun(C) Then
        O(J - 1) = " "
    Else
        O(J - 1) = C
    End If
Next
RplPun = Join(O, "")
End Function

Function RplVBar$(A)
RplVBar = Replace(A, "|", vbCrLf)
End Function

Function S1S2AyAdd(A() As S1S2, B() As S1S2) As S1S2()
Dim O() As S1S2
Dim J&
O = A
For J = 0 To UB(B)
    PushObj O, B(J)
Next
S1S2AyAdd = O
End Function

Function S1S2AyDic(A() As S1S2) As Dictionary
Dim J&, O As New Dictionary
For J = 0 To UB(A)
    O.Add A(J).S1, A(J).S2
Next
Set S1S2AyDic = O
End Function

Function S1S2AyFmt(A() As S1S2) As String()
Dim W1%: W1 = S1S2AyS1LinesWdt(A)
Dim W2%: W2 = S1S2AyS2LinesWdt(A)
Dim W%(1)
W(0) = W1
W(1) = W2
Dim H$: H = WdtAy_HdrLin(W)
S1S2AyFmt = S1S2AyLinesLinesLy(A, H, W1, W2)
End Function

Function S1S2AyLinesLinesLy(A() As S1S2, H$, W1%, W2%) As String()
Dim O$(), I&
Push O, H
For I = 0 To UB(A)
   PushAy O, S1S2Ly(A(I), W1, W2)
   Push O, H
Next
S1S2AyLinesLinesLy = O
End Function

Function S1S2AyS1LinesWdt%(A() As S1S2)
S1S2AyS1LinesWdt = LinesAyWdt(S1S2AySy1(A))
End Function

Function S1S2AyS2LinesWdt%(A() As S1S2)
S1S2AyS2LinesWdt = LinesAyWdt(S1S2AySy2(A))
End Function

Function S1S2AySy1(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S1
Next
S1S2AySy1 = O
End Function

Function S1S2AySy2(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S2
Next
S1S2AySy2 = O
End Function

Function S1S2Ly(A As S1S2, W1%, W2%) As String()
Dim S1$(), S2$()
S1 = SplitCrLf(A.S1)
S2 = SplitCrLf(A.S2)
Dim M%, J%, O$(), Lin$, A1$, A2$, U1%, U2%
    U1 = UB(S1)
    U2 = UB(S2)
    M = Max(U1, U2)
Dim Spc1$, Spc2$
    Spc1 = Space(W1)
    Spc2 = Space(W2)
For J = 0 To M
   If J > U1 Then
       A1 = Spc1
   Else
       A1 = StrAlignL(S1(J), W1)
   End If
   If J > U2 Then
       A2 = Spc2
   Else
       A2 = StrAlignL(S2(J), W2)
   End If
   Lin = "| " + A1 + " | " + A2 + " |"
   Push O, Lin
Next
S1S2Ly = O
End Function

Function SeedExpand$(QVbl$, Ny$())
Dim O$()
Dim Sy$(): Sy = SplitVBar(QVbl)
Dim J%, I
For J = 0 To UB(Ny)
    For Each I In Sy
       Push O, Replace(I, "?", Ny(J))
    Next
Next
SeedExpand = JnCrLf(O)
End Function

Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function

Function SplitDot(A) As String()
SplitDot = Split(A, ".")
End Function

Function SplitSsl(A) As String()
SplitSsl = Split(RplDblSpc(Trim(A)), " ")
End Function

Function SplitVBar(Vbl$) As String()
SplitVBar = Split(Vbl, "|")
End Function

Function SqWs(A, Optional WsNm$ = "Sheet1") As Worksheet
Dim A1 As Range: Set A1 = NewA1
SqRg A, A1
Set SqWs = RgWs(A1)
End Function

Function IsWhMdyAyVdt(A$()) As Boolean
Dim M
For Each M In A
    If Not AyHas(ShtMdyAy, M) Then Exit Function
Next
IsWhMdyAyVdt = True
End Function
Function CvWhMdy(WhMdy$) As String()
If WhMdy = "" Then Exit Function
Dim O$(), M
O = SslSy(WhMdy): If Not IsWhMdyAyVdt(O) Then Stop
If AyHas(O, "Pub") Then Push O, ""
CvWhMdy = O
End Function
Function CvWhMthKd(WhMthKd$) As String()
If WhMthKd = "" Then Exit Function
Dim O$(), K
O = SslSy(WhMthKd)
For Each K In O
    If Not AyHas(MthKdAy, K) Then Stop
Next
CvWhMthKd = O
End Function

Function IsPrpLin(A) As Boolean
IsPrpLin = LinMthKd(A) = "Property"
End Function

Function ItmAddAy(Itm, Ay)
Dim O, X
O = AyCln(Ay)
Push O, Itm
For Each X In AyNz(Ay)
    Push O, X
Next
ItmAddAy = O
End Function
Function MthFC(A As Mth) As FmCnt()
MthFC = SrcMthNmFC(MdBdyLy(A.Md), A.Nm)
End Function

Function FTIxFC(A As FTIx) As FmCnt
With A
    Set FTIxFC = FmCnt(.Fmix + 1, .Toix - .Fmix + 1)
End With
End Function

Function FTIxAyFC(A() As FTIx) As FmCnt()
FTIxAyFC = AyMapInto(A, "FTIxFC", FTIxAyFC)
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




Function ApLines$(ParamArray Ap())
Dim Av(): Av = Ap
ApLines = Join(AyRmvEmp(Av), vbCrLf)
End Function






Function IsRmkLin(A) As Boolean
IsRmkLin = FstChr(LTrim(A)) = "'"
End Function

Function MthEndLin$(MthLin$)
Dim A$
A = LinMthKd(MthLin): If A = "" Then Stop
MthEndLin = "End " & A
End Function

Function SslSy(Ssl) As String()
SslSy = Split(Trim(RplDblSpc(Ssl)), " ")
End Function
Function StrAlignL$(S$, W, Optional ErIFmnotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "StrAlignL"
Dim L%: L = Len(S)
If L > W Then
    If ErIFmnotEnoughWdt Then
        Stop
        'Er CSub, "Len({S)) > {W}", S, W
    End If
    If DoNotCut Then
        StrAlignL = S
        Exit Function
    End If
End If

If W >= L Then
    StrAlignL = S & Space(W - L)
    Exit Function
End If
If W > 2 Then
    StrAlignL = Left(S, W - 2) + ".."
    Exit Function
End If
StrAlignL = Left(S, W)
End Function

Function StrDup$(S, N%)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function StrNy(A) As String()
Dim O$: O = RplPun(A)
Dim O1$(): O1 = AyWhSingleEle(SslSy(O))
Dim O2$()
Dim J%
For J = 0 To UB(O1)
    If Not IsDigit(FstChr(O1(J))) Then Push O2, O1(J)
Next
StrNy = O2
End Function

Function SubStrCnt&(A, SubStr$)
Dim P&, O%, L%
L = Len(SubStr)
P = 1
While P > 0
    P = InStr(P, A, SubStr)
    If P > 0 Then O = O + 1: P = P + L
Wend
SubStrCnt = O
End Function

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
Function ShtMdyAy() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(0) = "Prv"
    O(1) = "Frd"
    O(2) = "Pub"
End If
ShtMdyAy = O
End Function

Function MthTyAy() As String()
Static O$(4), A As Boolean
If Not A Then
    A = True
    O(0) = "Property Get"
    O(1) = "Property Let"
    O(2) = "Property Set"
    O(3) = "Sub"
    O(4) = "Function"
End If
MthTyAy = O
End Function
Function MthKdAy() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(1) = "Sub"
    O(0) = "Function"
    O(2) = "Property"
End If
MthKdAy = O
End Function
Function MthShtTyAy() As String()
Static O$(4), A As Boolean
If Not A Then
    A = True
    O(0) = "Get"
    O(1) = "Let"
    O(2) = "Set"
    O(3) = "Sub"
    O(4) = "Fun"
End If
MthShtTyAy = O
End Function

Function SyOf_PrpSubFun() As String()
Static O$(2), A As Boolean
If Not A Then
    A = True
    O(0) = "Property"
    O(1) = "Sub"
    O(2) = "Function"
End If
SyOf_PrpSubFun = O
End Function

Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Function TmpFcsv$(Optional Fdr$, Optional Fnn$)
TmpFcsv = TmpFfn(".csv", Fdr, Fnn)
End Function

Function TmpFfn$(Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
If Fnn0 = "" Then
    Fnn = TmpNm
Else
    Fnn = Fnn0
End If
TmpFfn = TmpPth(Fdr) & Fnn & Ext
End Function

Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Function TmpFb$(Optional Fdr$, Optional Fnn$)
TmpFb = TmpFfn(".accdb", Fdr, Fnn)
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpPth$(Optional Fdr$)
Dim X$
   If Fdr <> "" Then
       X = Fdr & "\"
   End If
Dim O$
   O = TmpPthHom & X:   PthEns O
   O = O & TmpNm & "\": PthEns O
   PthEns O
TmpPth = O
End Function

Function TmpPthHom$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthHom = X
End Function

Function UB&(Ay)
UB = Sz(Ay) - 1
End Function
Function IsLinesAy(A) As Boolean
If Not IsSy(A) Then Exit Function
If Sz(A) = 0 Then Exit Function
Dim S
For Each S In A
    If IsLines(S) Then IsLinesAy = True: Exit Function
Next
End Function
Function IsLines(A) As Boolean
IsLines = True
If HasSubStr(A, vbCr) Then Exit Function
If HasSubStr(A, vbLf) Then Exit Function
IsLines = False
End Function
Function LinesAyLines$(A)
Dim W%
W = LinesAyWdt(A): If W = 0 Then Exit Function
LinesAyLines = Join(A, Space(W))
End Function

Function ObjToStr$(A)
If Not IsObject(A) Then Stop
On Error GoTo X
ObjToStr = A.ToStr: Exit Function
X: ObjToStr = QuoteSqBkt(TypeName(A))
End Function
Function QuoteSqBkt$(A)
QuoteSqBkt = "[" & A & "]"
End Function
Function IsSqBktQuoted(A) As Boolean
If FstChr(A) <> "[" Then Exit Function
IsSqBktQuoted = LasChr(A) = "]"
End Function
Function QuoteSqBktIfNeeded$(A)
Dim O$
Select Case True
Case IsSqBktQuoted(A): O = A
Case HasSpc(A): O = "[" & A & "]"
Case Else: O = A
End Select
QuoteSqBktIfNeeded = O
End Function
Function LvlSep$(Lvl%)
Select Case Lvl
Case 0: LvlSep = "."
Case 1: LvlSep = "-"
Case 2: LvlSep = "+"
Case 3: LvlSep = "="
Case 4: LvlSep = "*"
Case Else: LvlSep = Lvl
End Select
End Function

Private Sub ZZ_VarStr()
Dim A: A = Array(SslSy("sdf sdf df"), SslSy("sdf sdf"))
Debug.Print VarStr(A)
End Sub

Function VarStr$(A, Optional Lvl%)
Dim T$, S$, W%, I, O$(), Sep
Select Case True
Case IsPrim(A): VarStr = A
Case IsLinesAy(A): VarStr = LinesAyLines(A)
Case IsSy(A): VarStr = JnCrLf(A)
Case IsNothing(A): VarStr = "#Nothing"
Case IsEmpty(A): VarStr = "#Empty"
Case IsMissing(A): VarStr = "#Missing"
Case IsObject(A)
    VarStr = ObjToStr(A)
    T = TypeName(A)
    Select Case T
    Case "CodeModule"
        Dim M As CodeModule
        Set M = A
        VarStr = FmtQQ("*Md{?}", M.Parent.Name)
        Exit Function
    End Select
    VarStr = "*" & T
    Exit Function
Case IsArray(A)
    If Sz(A) = 0 Then Exit Function
    For Each I In A
        Push O, VarStr(I, Lvl + 1)
    Next
    W = LinesAyWdt(O)
    Sep = LvlSep(Lvl)
    VarStr = Join(O, vbCrLf & StrDup(Sep, W) & vbCrLf)
Case Else
End Select
End Function

Sub CurVbePjMdFmtBrw()
Brw VbePjMdFmt(CurVbe)
End Sub

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






Function WrpDryWdt(WrpDry(), WrpWdt%) As Integer() _
'WrpDry is dry having 1 or more wrpCol, which mean need wrapping.
'WrpWdt is for wrpCol _
'WrpCol is col with each cell being array
'if maxWdt of array-ele of wrpCol has wdt > WrpWdt, use that wdt
'otherwise use WrpWdt
If Sz(WrpDry) = 0 Then Exit Function
Dim J%, Col()
For J = 0 To DryNCol(WrpDry) - 1
    Col = DryCol(WrpDry, J)
    If IsArray(Col(0)) Then
        Push WrpDryWdt, AyWdt(AyFlat(Col))
    Else
        Push WrpDryWdt, AyWdt(Col)
    End If
Next
End Function

Function HasSpc(A) As Boolean
HasSpc = InStr(A, " ") > 0
End Function




Function UnEscCr$(A)
UnEscCr = Replace(A, "\r", vbCr)
End Function
Function UnEscLf$(A)
UnEscLf = Replace(A, "\n", vbCr)
End Function
Function UnEscCrLf$(A)
UnEscCrLf = UnEscLf(UnEscCr(A))
End Function
Function UnFmtss$(A)
UnFmtss = UnEscBackSlash(UnEscSqBkt(UnEscCrLf(A)))
End Function
Function UnEscBackSlash$(A)
UnEscBackSlash = Replace(A, "\\", "\")
End Function

Function UnEscSqBkt$(A)
UnEscSqBkt = Replace(A, Replace(A, "\o", "["), "\c", "]")
End Function
Function EscSqBkt$(A)
EscSqBkt = Replace(Replace(A, "[", "\o"), "]", "\c")
End Function
Function EscBackSlash$(A)
EscBackSlash = Replace(A, "\", "\\")
End Function
Function Fmtss$(A)
Fmtss = QuoteSqBktIfNeeded(EscSqBkt(EscCrLf(EscBackSlash(A))))
End Function

Function DrFmtssCell(A) As String()
Dim O$(), J&, X
O = AyReSz(O, A)
For Each X In AyNz(A)
    O(J) = Fmtss(X)
    J = J + 1
Next
DrFmtssCell = O
End Function

Function DrFmtssCellWrp(A, ColWdt%()) As String()
Dim X
For Each X In AyNz(A)
    PushIAy DrFmtssCellWrp, FmtssWrp(X, ColWdt)
Next
End Function
Function FmtssWrp(A, ColWdt%()) As String()

End Function
Function EscCrLf$(A)
EscCrLf = Replace(Replace(A, vbCr, "\r"), vbLf, "\n")
End Function

Function DrFmtss$(A, W%())
Dim U%, J%
U = UB(A)
If U = -1 Then Exit Function
ReDim O$(U)
For J = 0 To U - 1
    O(J) = AlignL(A(J), W%(J))
Next
O(U) = A(U)
DrFmtss = JnSpc(O)
End Function

Function SqRow(A, R%) As String()
Dim J%
For J = 1 To UBound(A, 2)
    Push SqRow, A(R, J)
Next
End Function

Function SqLy(A) As String()
Dim R%
For R = 1 To UBound(A, 1)
    Push SqLy, JnSpc(SqRow(A, R))
Next
End Function

Function WrpDrNRow%(WrpDr())
Dim Col, R%, M%
For Each Col In AyNz(WrpDr)
    M = Sz(Col)
    If M > R Then R = M
Next
WrpDrNRow = R
End Function

Function WrpDrSq(WrpDr()) As Variant()
Dim O(), R%, C%, NCol%, NRow%, Cell, Col, NColi%
NCol = Sz(WrpDr)
NRow = WrpDrNRow(WrpDr)
ReDim O(1 To NRow, 1 To NCol)
C = 0
For Each Col In WrpDr
    C = C + 1
    If IsArray(Col) Then
        NColi = Sz(Col)
        For R = 1 To NRow
            If R <= NColi Then
                O(R, C) = Col(R - 1)
            Else
                O(R, C) = ""
            End If
        Next
    Else
        O(1, C) = Col
        For R = 2 To NRow
            O(R, C) = ""
        Next
    End If
Next
WrpDrSq = O
End Function

Function SqAlign(Sq(), W%()) As Variant()
If UBound(Sq, 2) <> Sz(W) Then Stop
Dim C%, R%, Wdt%, O
O = Sq
For C = 1 To UBound(Sq, 2) - 1 ' The last column no need to align
    Wdt = W(C - 1)
    For R = 1 To UBound(Sq, 1)
        O(R, C) = AlignL(Sq(R, C), Wdt)
    Next
Next
SqAlign = O
End Function
Function WrpDrPad(WrpDr, W%()) As Variant() _
'Some Cell in WrpDr may be an array, wraping each element to cell if their width can fit its W%(?)
Dim J%, Cell, O()
O = WrpDr
For Each Cell In AyNz(O)
    If IsArray(Cell) Then
        O(J) = AyWrpPad(Cell, W(J))
    End If
    J = J + 1
Next
WrpDrPad = O
End Function
Function DrFmtssWrp(WrpDr, W%()) As String() _
'Each Itm of WrpDr may be an array.  So a DrFmt return Ly not string.
Dim Dr(): Dr = WrpDrPad(WrpDr, W)
Dim Sq(): Sq = WrpDrSq(Dr)
Dim Sq1(): Sq1 = SqAlign(Sq, W)
Dim Ly$(): Ly = SqLy(Sq1)
PushIAy DrFmtssWrp, Ly
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


Function VbeMthKy(A As Vbe, Optional IsWrap As Boolean) As String()
Dim O$(), I
For Each I In VbePjAy(A)
    PushAy O, PjMthKy(CvPj(I), IsWrap)
Next
VbeMthKy = O
End Function
Function MthNy() As String()
MthNy = CurVbeMthNy
End Function

Function MthNyWh(A As WhPjMth) As String()
MthNyWh = VbeMthNy(CurVbe, A)
End Function

Function CurVbeMthNy(Optional A As WhPjMth) As String()
CurVbeMthNy = VbeMthNy(CurVbe, A)
End Function

Function VbeMthNy(A As Vbe, Optional B As WhPjMth) As String()
Dim I, W As WhMdMth
Set W = WhPjMth_MdMth(B)
For Each I In AyNz(VbePjAy(A, WhPjMth_Nm(B)))
    PushIAy VbeMthNy, PjMthNy(CvPj(I), W)
Next
End Function

Function WhPjMth_Nm(A As WhPjMth) As WhNm
If IsNothing(A) Then Exit Function
Set WhPjMth_Nm = A.Pj
End Function

Function WhPjMth_MdMth(A As WhPjMth) As WhMdMth
If IsNothing(A) Then Exit Function
Set WhPjMth_MdMth = A.MdMth
End Function


Function VbePjAy(A As Vbe, Optional B As WhNm) As VBProject()
VbePjAy = ItrWhNmInto(A.VBProjects, B, VbePjAy)
End Function

Function VbePjNy(A As Vbe, Optional B As WhNm) As String()
VbePjNy = ItrNy(VbePjAy(A, B))
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

Function VblLines$(A)
VblLines = Replace(A, "|", vbCrLf)
End Function
Function WbHasWs(A As Workbook, WsNm$) As Boolean
WbHasWs = ItrHasNm(A.Sheets, WsNm)
End Function
Function WsSetNm(A As Worksheet, Nm$) As Worksheet
If Nm <> "" Then
    If Not WbHasWs(WsWb(A), Nm) Then A.Name = Nm
End If
Set WsSetNm = A
End Function
Function WbAddWs(A As Workbook, Optional WsNm$) As Worksheet
Set WbAddWs = WsSetNm(A.Sheets.Add(A.Sheets(1)), WsNm)
End Function

Function WbCn_TxtCn(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set WbCn_TxtCn = A.TextConnection
End Function

Function WbTxtCn(A As Workbook) As TextConnection
Dim N%: N = WbTxtCnCnt(A)
If N <> 1 Then
    Stop
    Exit Function
End If
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then
        Set WbTxtCn = C.TextConnection
        Exit Function
    End If
Next
ErImposs
End Function

Function WbTxtCnCnt%(A As Workbook)
Dim C As WorkbookConnection, Cnt%
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then Cnt = Cnt + 1
Next
WbTxtCnCnt = Cnt
End Function

Function WbTxtCnStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
If IsNothing(T) Then Exit Function
WbTxtCnStr = T.Connection
End Function

Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function

Function WdtAy_HdrLin$(A%())
Dim O$(), W
For Each W In A
    Push O, StrDup("-", W + 2)
Next
WdtAy_HdrLin = "|" + Join(O, "|") + "|"
End Function



Sub RgBdrTop(A As Range)
RgBdr A, xlEdgeTop
End Sub

Sub RgBdr(A As Range, Ix As XlBordersIndex, Optional Wgt As XlBorderWeight = xlMedium)
With A.Borders(Ix)
  .LineStyle = xlContinuous
  .Weight = Wgt
End With
End Sub
Function RgR(A As Range, R) As Range
Set RgR = RgRCRC(A, R, 1, R, RgNCol(A))
End Function
Function RgC(A As Range, C) As Range
Set RgC = RgRCRC(A, 1, RgNRow(A), 1, C)
End Function
Function RgNRow&(A As Range)
RgNRow = A.Rows.Count
End Function
Function RgNCol%(A As Range)
RgNCol = A.Columns.Count
End Function
Sub RgBdrAround(A As Range)
A.BorderAround XlLineStyle.xlContinuous, xlMedium
If A.Row > 1 Then RgBdrBottom RgR(A, 0)
If A.Column > 1 Then RgBdrRight RgC(A, 0)
RgBdrTop RgR(A, RgNRow(A) + 1)
RgBdrLeft RgC(A, RgNCol(A) + 1)
End Sub

Sub RgBdrBottom(A As Range)
RgBdr A, xlEdgeBottom
End Sub

Sub RgBdrInside(A As Range)
RgBdr A, xlInsideHorizontal
RgBdr A, xlInsideVertical
End Sub

Sub RgBdrLeft(A As Range)
RgBdr A, xlEdgeLeft
If A.Column > 1 Then
    RgBdr RgC(A, 0), xlEdgeRight
End If
End Sub

Sub RgBdrRight(A As Range)
RgBdr A, xlEdgeRight
If A.Column < MaxCol Then
    RgBdr RgC(A, A.Column + 1), xlEdgeLeft
End If
End Sub





Function Xls() As Excel.Application
Static Y As Excel.Application
On Error GoTo X
Dim A$: A = Y.Name
Set Xls = Y
Exit Function
X:
Set Y = New Excel.Application
Set Xls = Y
End Function

Function XlsHasAddInFn(A As Excel.Application, AddInFn) As Boolean
Dim I As Excel.AddIn
Dim N$: N = UCase(AddInFn)
For Each I In A.AddIns
    If UCase(I.Name) = N Then XlsHasAddInFn = True: Exit Function
Next
End Function

Sub Asg(V, OV)
If IsObject(V) Then
   Set OV = V
Else
   OV = V
End If
End Sub

Sub Ass(A As Boolean)
Debug.Print A
End Sub

Sub D(A)
Select Case True
Case IsStr(A): Debug.Print A
Case IsNumeric(A): Debug.Print A
Case IsArray(A)
    Dim X
    For Each X In AyNz(A)
        D X
    Next
Case IsEmpty(A): Debug.Print "*Empty"
Case IsNothing(A): Debug.Print "*Nothing"
Case IsObject(A): Debug.Print ObjToStr(A)
Case Else
Stop
End Select
End Sub






Sub CmpRmv(A As VBComponent)
A.Collection.Remove A
End Sub

Sub DDNmBrkAsg(A, O1$, O2$, O3$)
Dim Ay$(): Ay = Split(A, ".")
Select Case Sz(Ay)
Case 1: O1 = "":    O2 = "":    O3 = Ay(0)
Case 2: O1 = "":    O2 = Ay(0): O3 = Ay(1)
Case 3: O1 = Ay(0): O2 = Ay(1): O3 = Ay(2)
Case Else: Stop
End Select
End Sub
Function TyNm$(A)
TyNm = TypeName(A)
End Function
Sub DicTyBrw(A As Dictionary)
DicBrw DicTy(A)
End Sub
Function DicTy(A As Dictionary) As Dictionary
Set DicTy = DicMap(A, "TyNm")
End Function
Sub DicWsBrw(A As Dictionary)
WsVis DicWs(A)
End Sub
Sub DicBrw(A As Dictionary)
Brw DicLy(A)
End Sub
Function DicLy(A As Dictionary) As String()
DicLy = S1S2AyFmt(DicS1S2Ay(A))
End Function

Sub DupMthFNy_ShwNotDupMsg(A$(), MthNm)
Select Case Sz(A)
Case 0: Debug.Print FmtQQ("DupMthFNy_ShwNotDupMsg: no such Fun(?) in CurVbe", MthNm)
Case 1
    Dim B As S1S2: Set B = Brk(A(0), ":")
    If B.S1 <> MthNm Then Stop
    Debug.Print FmtQQ("DupMthFNy_ShwNotDupMsg: Fun(?) in Md(?) does not have dup-Fun", B.S1, B.S2)
End Select
End Sub

Sub ErImposs()
Stop ' Impossible
End Sub



Sub FtBrw(A)
Shell "code.cmd """ & A & """", vbHide
'Shell "notepad.exe """ & A & """", vbMaximizedFocus
End Sub

Sub FtRmvFst4Lines(FT$)
Dim A$: A = Fso.GetFile(FT).OpenAsTextStream.ReadAll
Dim B$: B = Left(A, 55)
Dim C$: C = Mid(A, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If B <> B1 Then Stop
Fso.CreateTextFile(FT, True).Write C
End Sub

Sub FunFNm_BrkAsg(A$, OFunNm$, OPjNm$, OMdNm$)
With Brk(A, ":")
    OFunNm = .S1
    With Brk(.S2, ".")
        OPjNm = .S1
        OMdNm = .S2
    End With
End With
End Sub

Sub FxaNm_Crt(A)
FxaCrt FxaNm_Fxa(A)
End Sub

Sub FxaCrt(A)
If FfnIsExist(A) Then
    Debug.Print FmtQQ("FxaCrt: Fxa(?) is already exist", A)
    Exit Sub
End If
If XlsHasAddInFn(CurXls, FfnFn(A)) Then Stop: Exit Sub
Dim O As Workbook
Set O = CurXls.Workbooks.Add
O.SaveAs A, XlFileFormat.xlOpenXMLAddIn
O.Close
Dim AddIn As AddIn: Set AddIn = CurXls.AddIns.Add(A)
AddIn.Installed = True
Dim Pj As VBProject
Set Pj = VbePjFfn_Pj(CurVbe, A)
Pj.Name = FfnFnn(A)
PjSav Pj
End Sub
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
Function XlsAddIn(A As Excel.Application, FxaNm) As Excel.AddIn
Dim I As Excel.AddIn
For Each I In A.AddIns
    If StrIsEq(I.Name, FxaNm & ".xlam") Then Set XlsAddIn = I
Next
End Function
Function StrIsEq(A, B) As Boolean
StrIsEq = StrComp(A, B, vbTextCompare) = 0
End Function
Sub ItrDoSub(A, SubNm$)
Dim I
For Each I In A
    CallByName A, SubNm, VbMethod
Next
End Sub

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

Sub MthBrkAsg(A As Mth, OMdy$, OMthTy$)
Dim L$: L = MthLin(A)
OMdy = TakMdy(L)
OMthTy = LinMthTy(L)
End Sub

Sub MthGo(A As Mth)
MdGoMayLCC A.Md, MthMayLCC(A)
End Sub

Function JnSpc$(A)
JnSpc = Join(A, " ")
End Function

Function DicWs(A As Dictionary) As Worksheet
Set DicWs = S1S2AyWs(DicS1S2Ay(A))
End Function

Function ItrSy(A) As String()
Dim O$(), I, J&
If A.Count = 0 Then Exit Function
ReDim O(A.Count - 1)
For Each I In A
    O(J) = I
    J = J + 1
Next
ItrSy = O
End Function

Function DicStrKy(A As Dictionary) As String()
DicStrKy = AySy(A.Keys)
End Function

Function DicMaxValSz%(A As Dictionary)
'MthDic is DicOf_MthNm_zz_MthLinesAy
'MaxMthCnt is max-of-#-of-method per MthNm
Dim O%, K
For Each K In A.Keys
    O = Max(O, Sz(A(K)))
Next
DicMaxValSz = O
End Function

Function MthCpyPrm_Cpy(A As MthCpyPrm)
MthCpy A.SrcMth, A.ToMd
End Function

Function DicAyAdd(A() As Dictionary) As Dictionary
Dim O As New Dictionary, D
For Each D In A
    PushDic O, CvDic(D)
Next
Set DicAyAdd = O
End Function
Sub PushDic(O As Dictionary, A As Dictionary)
Dim K
For Each K In A.Keys
    If O.Exists(K) Then Stop
    O.Add K, A(K)
Next
End Sub

Function ApSy(ParamArray Ap()) As String()
Dim Av(): Av = Ap
ApSy = AySy(Av)
End Function

Function IsSyAy(A) As Boolean
If Not IsArray(A) Then Exit Function
If Sz(A) = 0 Then IsSyAy = True: Exit Function
Dim I
For Each I In A
    If Not IsSy(I) Then Exit Function
Next
IsSyAy = True
End Function

Function S1S2AySq(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Variant()
If Sz(A) = 0 Then Exit Function
Dim O(), I, R&
ReDim O(1 To Sz(A), 1 To 2)
R = 2
O(1, 1) = Nm1
O(1, 2) = Nm2
For Each I In AyNz(A)
    With CvS1S2(I)
        O(R, 1) = .S1
        O(R, 2) = .S2
        R = R + 1
    End With
Next
S1S2AySq = O
End Function


Function S1S2AyWs(A() As S1S2, Optional Nm1$ = "S1", Optional Nm2$ = "S2") As Worksheet
Set S1S2AyWs = SqWs(S1S2AySq(A, Nm1, Nm2))
End Function

Function IsFmCntInOrd(A() As FmCnt) As Boolean
Dim J%
For J = 0 To UB(A) - 1
    With A(J)
      If .FmLno + .Cnt > A(J + 1).FmLno Then Exit Function
    End With
Next
IsFmCntInOrd = True
End Function
Sub MdRmvFC(A As CodeModule, B() As FmCnt)
If Not IsFmCntInOrd(B) Then Stop
Dim J%
For J = UB(B) To 0 Step -1
    With B(J)
        A.DeleteLines .FmLno, .Cnt
    End With
Next
End Sub
Sub MthRmv(A As Mth, Optional IsSilent As Boolean)
Dim X() As FmCnt: X = MthRmkFC(A)
MdRmvFC A.Md, X
If Not IsSilent Then
    Debug.Print FmtQQ("MthRmv: Mth(?) of LinCnt(?) is deleted", MthDNm(A), FmCntAyLinCnt(X))
End If
End Sub

Sub MthRpl(A As Mth, By$)
MthRmv A
MdAppLines A.Md, By
End Sub

Sub OyDo(Oy, DoFun$)
Dim O
For Each O In Oy
    Excel.Run DoFun, O ' DoFunNm cannot be like a Excel.Address (eg, A1, XX1)
Next
End Sub

Sub PjAddCls(A As VBProject, Nm$)
PjAddMbr A, Nm, vbext_ct_ClassModule
End Sub

Sub PjAddMbr(A As VBProject, Nm$, Ty As vbext_ComponentType, Optional IsGoMbr As Boolean)
If PjHasCmp(A, Nm) Then
    MsgBox FmtQQ("Cmp(?) exist in CurPj(?)", Nm, CurPjNm), , "M_A.ZAddMbr"
    Exit Sub
End If
Dim Cmp As VBComponent
Set Cmp = A.VBComponents.Add(Ty)
Cmp.Name = Nm
Cmp.CodeModule.InsertLines 1, "Option Explicit"
'If IsGoMbr Then ShwMbr Nm
End Sub
Private Sub ZZ_PjAddRf()
PjAddRf Pj("QXls"), "QDta"
End Sub
Sub PjRmvRf(A As VBProject, RfNy0$)
AyDoPX DftNy(RfNy0), "PjRmvRf__X", A
PjSav A
End Sub
Sub PjAddRf(A As VBProject, RfNy0$)
AyDoPX DftNy(RfNy0), "PjAddRf__X", A
PjSav A
End Sub
Private Sub PjAddRf__X(A As VBProject, RfNm$)
If PjHasRfNm(A, RfNm) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfNm(?)", A.Name, RfNm)
    Exit Sub
End If
Dim RfFfn$: RfFfn = PjRfNm_RfFfn(A, RfNm)
If PjHasRfFfn(A, RfFfn) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfFfnNm(?)", A.Name, RfFfn)
    Exit Sub
End If
A.References.AddFromFile RfFfn
End Sub
Private Sub PjRmvRf__X(A As VBProject, RfNm$)
If PjHasRfNm(A, RfNm) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfNm(?)", A.Name, RfNm)
    Exit Sub
End If
Dim RfFfn$: RfFfn = PjRfNm_RfFfn(A, RfNm)
If PjHasRfFfn(A, RfFfn) Then
    Debug.Print FmtQQ("PjAddRf: Pj(?) already has RfFfnNm(?)", A.Name, RfFfn)
    Exit Sub
End If
A.References.AddFromFile RfFfn
End Sub
Sub ZZ_VbeWsFunNmzDupLines()
WsVis VbeWsFunNmzDupLines(CurVbe)
End Sub

Function VbeWsFunNmzDupLines(A As Vbe) As Worksheet
Set VbeWsFunNmzDupLines = DrsWs(VbeDrsFunNmzDupLines(A))
End Function

Function VbeDrsFunNmzDupLines(A As Vbe) As Drs
'Nm AllLinesEq N Lines01....
'Dim Drs As Drs
'Set Drs = VbeFun12Drs(A)
'Set Drs = DrsGpFlat(Drs, "Nm", "Lines")
'Set Drs = DrsWhColGt(Drs, "N", 2)
'Set Drs = DrsInsColBef(Drs, "N", "AllLinesEq")
'Set Drs = VbeDrsFunNmzDupLines__1(Drs)
'Set VbeDrsFunNmzDupLines = Drs
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
Private Sub Z_LblSeqAy()
Dim Act$(), A, N%, Exp$()
A = "Lbl"
N = 10
Exp = SslSy("Lbl01 Lbl02 Lbl03 Lbl04 Lbl05 Lbl06 Lbl07 Lbl08 Lbl09 Lbl10")
Act = LblSeqAy(A, N)
Ass IsEqAy(Act, Exp)
End Sub

Function LblSeqSsl$(A, N%)
LblSeqSsl = Join(LblSeqAy(A, N), " ")
End Function

Function LblSeqAy(A, N%) As String()
Dim O$(), J%, F$, L%
L = Len(N)
F = StrDup("0", L)
ReDim O(N - 1)
For J = 1 To N
    O(J - 1) = A & Format(J, F)
Next
LblSeqAy = O
End Function
Function EmpLngAy() As Long()
End Function
Function AyAddFunCol(A, FunNm$) As Variant()
Dim X
For Each X In AyNz(A)
    PushI AyAddFunCol, Array(X, Run(FunNm, X))
Next
End Function

Sub Sts()
Dim A$(), B$(), C()
A = SrcMthNy(MdBdyLy(Md("QTool.AX")))
B = AySrt(AyMapSy(A, "MthPfx"))
C = AyGpCntDryWhDup(B)
D DryFmtss(C)
D "Cnt=" & Sz(C)
End Sub
Private Sub ZZ_DrsGpDic()
Dim Act As Dictionary, Dry(), Dr1, Dr2, Dr3
Dr1 = Array("A", , 1)
Dr2 = Array("A", , 2)
Dr3 = Array("B", , 3)
Dry = Array(Dr1, Dr2, Dr3)
Set Act = DryGpDic(Dry, 0, 2)
Ass Act.Count = 2
Ass IsEqAy(Act("A"), Array(1, 2))
Ass IsEqAy(Act("B"), Array(3))
Stop
End Sub

Function DrsGpFlat(A As Drs, K$, G$) As Drs
Dim Fny0$, Dry(), S$, N%, Ix%()
Ix = AyIxAyI(A.Fny, Array(K, G))
Dry = DryGpFlat(A.Dry, Ix(0), Ix(1))
N = DryNCol(Dry) - 2
S = LblSeqSsl(G, N)
Fny0 = FmtQQ("? N ?", K, S)
Set DrsGpFlat = Drs(Fny0, Dry)
End Function

Sub ZZ_PjCompile()
PjCompile CurPj
End Sub
Sub VbeCompile(A As Vbe)
ItrDo A.VBProjects, "PjCompile"
End Sub
Sub PjCompile(A As VBProject)
PjGo A
AssCompileBtn PjNm(A)
With CompileBtn
    If .Enabled Then
        .Execute
        Debug.Print PjNm(A), "<--- Compiled"
    Else
        Debug.Print PjNm(A), "already Compiled"
    End If
End With
TileVBtn.Execute
SavBtn.Execute
End Sub

Sub PjCrt_Fxa(A As VBProject, FxaNm$)
Dim F$
F = FxaNm_Fxa(FxaNm)
End Sub

Function PjEnsCls(A As VBProject, ClsNm$) As CodeModule
Set PjEnsCls = PjEnsCmp(A, ClsNm, vbext_ct_ClassModule).CodeModule
End Function

Function PjEnsCmp(A As VBProject, Nm$, Ty As vbext_ComponentType) As VBComponent
If Not PjHasCmp(A, Nm) Then
    Dim Cmp As VBComponent
    Set Cmp = A.VBComponents.Add(Ty)
    Cmp.Name = Nm
    Cmp.CodeModule.AddFromString "Option Explicit"
    Debug.Print FmtQQ("PjEnsCmp: Md(?) of Ty(?) is added in Pj(?) <===================================", Nm, CmpTy_Nm(Ty), A.Name)
End If
Set PjEnsCmp = A.VBComponents(Nm)
End Function

Function PjEnsMd(A As VBProject, MdNm$) As CodeModule
Set PjEnsMd = PjEnsCmp(A, MdNm, vbext_ct_StdModule).CodeModule
End Function

Sub PjExport(A As VBProject)
Dim P$: P = PjSrcPth(A)
If P = "" Then
    Debug.Print FmtQQ("PjExport: Pj(?) does not have FileName", A.Name)
    Exit Sub
End If
PthClrFil P 'Clr SrcPth ---
FfnCpyToPth A.Filename, P, OvrWrt:=True
Dim I, Ay() As CodeModule
Ay = PjMdAy(A)
If Sz(Ay) = 0 Then Exit Sub
For Each I In Ay
    MdExport CvMd(I)  'Exp each md --
Next
AyWrt PjRfLy(A), PjRfCfgFfn(A) 'Exp rf -----
End Sub
Sub PjGo(A As VBProject)
ClsWin
Dim Md As CodeModule
Set Md = PjFstMbr(A)
If IsNothing(Md) Then
    Exit Sub
End If
Md.CodePane.Show
TileVBtn.Execute
DoEvents
End Sub
Function PjTim(A As VBProject) As Date
PjTim = FfnTim(PjFfn(A))
End Function


Function PjFn$(A As VBProject)
PjFn = FfnFn(PjFfn(A))
End Function
Private Sub ZZ_PjSav()
PjSav CurPj
End Sub
Sub VbeSav(A As Vbe)
ItrDo A.VBProjects, "PjSav"
End Sub

Private Sub ZZ_VbeDmpIsSaved()
VbeDmpIsSaved CurVbe
End Sub
Sub VbeDmpIsSaved(A As Vbe)
Dim I As VBProject
For Each I In A.VBProjects
    Debug.Print I.Saved, I.BuildFileName
Next
End Sub
Function ItrPrpAy(A, PrpNm)
ItrPrpAy = ItrPrpAyInto(A, PrpNm, EmpAy)
End Function
Function ItrPrpAyInto(A, PrpNm, OInto)
Dim O: O = OInto: Erase O
Dim I
For Each I In A
    Push O, ObjPrp(I, PrpNm)
Next
ItrPrpAyInto = O
End Function
Sub ItrDo(A, DoFunNm$)
Dim I
For Each I In A
    Run DoFunNm, I
Next
End Sub
Sub PjAddClsFmPj(A As VBProject, FmPj As VBProject, ClsNy0)
Dim I, ClsNy$(), ClsAy() As CodeModule
ClsNy = DftNy(ClsNy0)
'ClsAy = PjMd(
For Each I In A
    MdCpy CvMd(I), A
Next
End Sub

Sub PjSav(A As VBProject)
If FstChr(PjNm(A)) <> "Q" Then
    Exit Sub
End If
If A.Saved Then
    Debug.Print FmtQQ("PjSav: Pj(?) is already saved", A.Name)
    Exit Sub
End If
Dim Fn$: Fn = PjFn(A)
If Fn = "" Then
    Debug.Print FmtQQ("PjSav: Pj(?) needs saved first", A.Name)
    Exit Sub
End If
PjGo A
If ObjPtr(CurPj) <> ObjPtr(A) Then Stop: Exit Sub
Dim B As CommandBarButton: Set B = SavBtn
If Not StrIsEq(B.Caption, "&Save " & Fn) Then Stop
B.Execute
Debug.Print FmtQQ("PjSav: Pj(?) is not sure if saved <---------------", A.Name)
End Sub

Sub PjSrcPthBrw(A As VBProject)
PthBrw PjSrcPth(A)
End Sub

Sub PjSrt(A As VBProject)
Dim I
Dim Ny$(): Ny = AySrt(PjStdClsNy(A))
If Sz(Ny) = 0 Then Exit Sub
For Each I In Ny
    MdSrt PjMd(A, I)
Next
End Sub

Sub Pj_Gen_TstClass(A As VBProject)
If PjHasCmp(A, "Tst") Then
    CmpRmv PjCmp(A, "Tst")
End If
PjAddCls A, "Tst"
PjMd(A, "Tst").AddFromString Pj_TstClass_Bdy(A)
End Sub

Sub Pj_Gen_TstSub(A As VBProject)
Dim Ny$(): Ny = PjStdClsNy(A)
Dim N, M As CodeModule
For Each N In Ny
    Set M = A.VBComponents(N).CodeModule
    Md_Gen_TstSub M
Next
End Sub



Sub PushI(O, M)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
O(N) = M
End Sub
Sub PushIAy(O, MAy)
Dim M
For Each M In AyNz(MAy)
    PushI O, M
Next
End Sub
Sub Push(O, M)
Dim N&
    N = Sz(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub

Sub PushAy(OAy, Ay)
If Sz(Ay) = 0 Then Exit Sub
Dim I
For Each I In Ay
    Push OAy, I
Next
End Sub

Sub PushAyNoDup(OAy, Ay)
If Sz(Ay) = 0 Then Exit Sub
Dim I
For Each I In Ay
    PushNoDup OAy, I
Next
End Sub

Sub PushNoDup(O, M)
If Not AyHas(O, M) Then PushI O, M
End Sub
Sub PushNonBlankStr(O, M$)
If M = "" Then Exit Sub
PushI O, M
End Sub
Sub PushNonEmp(O, M)
If IsEmp(M) Then Exit Sub
Push O, M
End Sub
Sub PushISomSz(OAy, IAy)
If Sz(IAy) = 0 Then Exit Sub
PushI OAy, IAy
End Sub

Sub PushAyNonZSz(OAy, Ay)
If Sz(Ay) = 0 Then Exit Sub
PushIAy OAy, Ay
End Sub

Sub PushObj(O, M)
If Not IsObject(M) Then Stop
Dim N&
    N = Sz(O)
ReDim Preserve O(N)
Set O(N) = M
End Sub

Sub PushObjAy(O, Oy)
If Sz(Oy) = 0 Then Exit Sub
Dim I
For Each I In Oy
    PushObj O, I
Next
End Sub

Function RgVis(A As Range) As Range
A.Application.Visible = True
Set RgVis = A
End Function

Sub S1S2AyBrw(A() As S1S2)
Brw S1S2AyFmt(A)
End Sub

Sub SqSetRow(OSq, R&, Dr)
Dim J%
For J = 0 To UB(Dr)
    OSq(R, J + 1) = Dr(J)
Next
End Sub
Function StrLikItr(A, LikItr As Collection) As Boolean
Dim I
For Each I In LikItr
    If A Like I Then StrLikItr = True
Next
End Function

Sub StrBrw(A)
Dim T$:
T = TmpFt
StrWrt A, T
Shell FmtQQ("code.cmd ""?""", T), vbMaximizedFocus
'Shell FmtQQ("notepad.exe ""?""", T), vbMaximizedFocus
End Sub

Sub StrWrt(A, FT$, Optional IsNotOvrWrt As Boolean)
Fso.CreateTextFile(FT, Overwrite:=Not IsNotOvrWrt).Write A
End Sub

Sub CurVbeExport()
VbeExport CurVbe
End Sub

Sub Export()
CurVbeExport
End Sub

Sub VbeExport(A As Vbe)
OyDo VbePjAy(A), "PjExport"
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

Function WbSavAs(A As Workbook, Fx) As Workbook
A.SaveAs Fx
Set WbSavAs = A
End Function

Function WbRfh(A As Workbook) As Workbook
Dim Ws As Worksheet
For Each Ws In A.Worksheets
    WsRfh Ws
Next
Dim Pc As PivotCache
For Each Pc In A.PivotCaches
    Pc.MissingItemsLimit = xlMissingItemsNone
    Pc.Refresh
Next
Set WbRfh = A
End Function

Sub WbSetFcsv(A As Workbook, Fcsv$)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub


Sub XlsAddFxaNm(A As Excel.Application, FxaNm$)
Dim F$: F = FxaNm_Fxa(FxaNm)
If F = "" Then Exit Sub
A.AddIns.Add FxaNm_Fxa(FxaNm)
End Sub

Sub XlsVis(A As Excel.Application)
If Not A.Visible Then A.Visible = True
End Sub

Private Function DupMthFNyGp_CmpLy__1Hdr(OIx%, MthNm$, Cnt%) As String()
Dim O$(1)
O(0) = "================================================================"
Dim A$
    If OIx >= 0 Then A = FmtQQ("#DupMthNo(?) ", OIx): OIx = OIx + 1
O(1) = A + FmtQQ("DupMthNm(?) Cnt(?)", MthNm, Cnt)
DupMthFNyGp_CmpLy__1Hdr = O
End Function

Private Function DupMthFNyGp_CmpLy__2Sam(InclSam As Boolean, OSam%, DupMthFNyGp, LinesAy$()) As String()
If Not InclSam Then Exit Function
'{DupMthFNyGp} & {LinesAy} have same # of element
Dim O$()
Dim D$(): D = AyWhDup(LinesAy)
Dim J%, X$()
For J = 0 To UB(D)
    X = DupMthFNyGp_CmpLy__2Sam1(OSam, D(J), DupMthFNyGp, LinesAy)
    PushAy O, X
Next
DupMthFNyGp_CmpLy__2Sam = O
End Function

Private Function DupMthFNyGp_CmpLy__2Sam1(OSam%, Lines$, DupMthFNyGp, LinesAy$()) As String()
Dim A1$()
    If OSam > 0 Then
        Push A1, FmtQQ("#Sam(?) ", OSam)
        OSam = OSam + 1
    End If
Dim A2$()
    Dim J%
    For J = 0 To UB(LinesAy)
        If LinesAy(J) = Lines Then
            Push A2, "Shw """ & DupMthFNyGp(J) & """"
        End If
    Next
Dim A3$()
    A3 = LinesBoxLy(Lines)
DupMthFNyGp_CmpLy__2Sam1 = AyAddAp(A1, A2, A3)
End Function

Private Function DupMthFNyGp_CmpLy__3Syn(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim B$()
    Dim J%, I%
    Dim Lines
    For Each Lines In UniqLinesAy
        For I = 0 To UB(FunFNyGp)
            If Lines = LinesAy(I) Then
                Push B, FunFNyGp(I)
                Exit For
            End If
        Next
    Next
DupMthFNyGp_CmpLy__3Syn = AyMapPXSy(B, "FmtQQ", "Sync_Fun ""?""")
End Function

Private Function DupMthFNyGp_CmpLy__4Cmp(UniqLinesAy$(), LinesAy$(), FunFNyGp) As String()
If Sz(UniqLinesAy) <= 1 Then Exit Function
Dim L2$() ' = From L1 with each element with MdDNm added in front
    ReDim L2(UB(UniqLinesAy))
    Dim Fnd As Boolean, DNm$, J%, Lines$, I%
    For J = 0 To UB(UniqLinesAy)
        Lines = UniqLinesAy(J)
        Fnd = False
        For I = 0 To UB(LinesAy)
            If LinesAy(I) = Lines Then
                DNm = FunFNyGp(I)
                L2(J) = DNm & vbCrLf & StrDup("-", Len(DNm)) & vbCrLf & Lines
                Fnd = False
                GoTo Nxt
            End If
        Next
        Stop
Nxt:
    Next
DupMthFNyGp_CmpLy__4Cmp = LinesAyFmt(L2)
End Function


Private Property Get ZZSrc() As String()
ZZSrc = MdSrc(CurMd)
End Property

Private Sub Z_MdEndTrim()
Dim M As CodeModule: Set M = Md("ZZModule")
MdAppLines M, "  "
MdAppLines M, "  "
MdAppLines M, "  "
MdAppLines M, "  "
MdEndTrim M, ShwMsg:=True
Ass M.CountOfLines = 15
End Sub

Private Sub Z_MthFC()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
Dim Act() As FmCnt: Act = MthFC(M)
Ass Sz(Act) = 2
Ass Act(0).FmLno = 5
Ass Act(0).Cnt = 7
Ass Act(1).FmLno = 13
Ass Act(1).Cnt = 15
End Sub

Function CurPjEnsMd(MdNm$) As CodeModule
Set CurPjEnsMd = PjEnsMd(CurPj, MdNm)
End Function

Sub PjDltMd(A As VBProject, MdNm$)
If Not PjHasMd(A, MdNm) Then Exit Sub
A.VBComponents.Remove A.VBComponents(MdNm)
End Sub

Sub CurPjDltMd(MdNm$)
PjDltMd CurPj, MdNm
End Sub

Private Sub Z_MthRmv()
Const N$ = "ZZModule"
Dim M As CodeModule
Dim M1 As Mth, M2 As Mth
GoSub Crt
Set M = Md(N)
Set M1 = Mth(M, "ZZRmv1")
Set M2 = Mth(M, "ZZRmv2")
MthRmv M1
MthRmv M2
MdEndTrim M
If M.CountOfLines <> 1 Then MsgBox M.CountOfLines
MdDlt M
Exit Sub
Crt:
    CurPjDltMd N
    Set M = CurPjEnsMd(N)
    MdAppLines M, RplVBar("Property Get ZZRmv1()||End Property||Function ZZRmv2()|End Function||'|Property Let ZZRmv1(V)|End Property")
    Return
End Sub

Private Sub Z_WbSetFcsv()
Dim Wb As Workbook
Set Wb = FxWb(VbeMthFx)
Debug.Print WbTxtCnStr(Wb)
WbSetFcsv Wb, "C:\ABC.CSV"
Ass WbTxtCnStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub

Private Sub Z_WbTxtCnCnt()
Dim O As Workbook: Set O = FxWb(VbeMthFx)
Ass WbTxtCnCnt(O) = 1
O.Close
End Sub

Private Sub ZZ_LinesAyFmt()
Dim A$()
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf|sdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdf|skldfjdf|lskdf|slkdjf|sdlf||")
Push A, RplVBar("ksdjlfdf|sdklfjsdfdsfdsf|skldsdffjdf")
D LinesAyFmt(A)
End Sub

Private Sub ZZ_MthNmSrtKey()
Dim Ay1$(): Ay1 = SrcMthNy(CurSrc)
Dim Ay2$(): Ay2 = AyMapSy(Ay1, "MthNmSrtKey")
S1S2AyBrw AyabS1S2Ay(Ay2, Ay1)
End Sub

Private Sub ZZ_MthNmSrtKey_1()
Const A$ = "YYA.Fun."
Debug.Print MthNmSrtKey(A)
End Sub

Private Sub ZZ_SrcDclLinCnt()
Dim B1$(): B1 = CurSrc
Dim B2$(): B2 = SrcSrtedLy(B1)
Dim A1%: A1 = SrcDclLinCnt(B1)
Dim A2%: A2 = SrcDclLinCnt(SrcSrtedLy(B1))
End Sub

Private Sub ZZ_MdSrtedLines()
StrBrw MdSrtedLines(CurMd)
End Sub

Private Sub ZZ_VbeMthNy()
Brw VbeMthNy(CurVbe)
End Sub

Function VbeMthPfx(A As Vbe) As String()

End Function
Private Sub ZZ_VbeFunPfx()
D VbeMthPfx(CurVbe)
End Sub

Private Sub ZZ_XlsAddFxaNm()
XlsAddFxaNm Application, "QIde0"
End Sub

Function DftFun(FunDNm0$) As Mth
If FunDNm0 = "" Then
    Dim M As Mth
    Set M = CurMth
    If IsFun(M) Then
        Set DftFun = M
    End If
Else
End If
Stop '
End Function

Function IsMthDNm(Nm) As Boolean
IsMthDNm = Sz(Split(Nm, ".")) = 3
End Function

Function IsMthFNm(Nm) As Boolean
Dim P%: P = InStr(Nm, ":"): If P = 0 Then Exit Function
IsMthFNm = InStr(Nm, ".") > P
End Function

Function FxBrw(A$)
WbVis FxWb(A)
End Function
Function InsDrApSqy(T, Fny0, ParamArray DrAp()) As String()
Dim Dr, Av(), Fny$()
Fny = CvNy(Fny0)
Av = DrAp
For Each Dr In Av
    PushI InsDrApSqy, InsDrSql(T, Fny, Dr)
Next
End Function
Function InsDrSql$(T, Fny0, Dr)
InsDrSql = FmtQQ("Insert into [?] (?) values(?)", T, JnComma(CvNy(Fny0)), JnComma(SqlValAyQuote(Dr)))
End Function
Function DbtDry(A As Database, T) As Variant()
DbtDry = RsDry(DbtRs(A, T))
End Function
Function DbtRs(A As Database, T) As DAO.Recordset
Set DbtRs = A.OpenRecordset(SelTblSql(T))
End Function
Function RsDry(A As DAO.Recordset) As Variant()
With A
    While Not .EOF
        PushI RsDry, FldsDr(A.Fields)
        .MoveNext
    Wend
End With
End Function

Function FldsDr(A As DAO.Fields) As Variant()
Dim F As DAO.Field
For Each F In A
    PushI FldsDr, F.Value
Next
End Function

Function SelTblSql$(T)
SelTblSql = "Select * from [" & T & "]"
End Function
Sub AAAAA()
Z_DrsInsUpdDbt
End Sub
Sub Z_DrsInsUpdDbt()
Dim Db As Database, T, A As Drs, TFb$
    TFb = TmpFb("Tst", "DrsInsUpdDbt")
    Set Db = FbCrt(TFb)
T = "Tmp"
Db.Execute "Create Table Tmp (A Int, B Int, C Int)"
Db.Execute CrtSkSql("Tmp", "A")
DbSqyRun Db, InsDrApSqy("Tmp", "A B C", Array(1, 3, 4), Array(3, 4, 5))
Set A = Drs("A B C", CvAy(Array(Array(1, 2, 3), Array(2, 3, 4))))

Ept = Array(Array(1&, 2&, 3&), Array(2&, 3&, 4&), Array(3&, 4&, 5&))
GoSub Tst
Db.Close
Kill TFb
Exit Sub
Tst:
    DrsInsUpdDbt A, Db, T
    Act = DbtDry(Db, T)
    C
    Return
End Sub

Sub DrsInsUpdDbt(A As Drs, Db As Database, T)
GoSub X
Dim Ins As Drs, Upd As Drs: GoSub X_Ins_Upd
DrsInsDbt Ins, Db, T
DrsUpdDbt Upd, Db, T
Exit Sub
X:
    Dim Sk$()
    Dim SkIxAy&()
    Dim Fny$()
    Dim Dry()
    Sk = DbtSk(Db, T)
    Fny = A.Fny
    SkIxAy = AyIxAy(Fny, Sk)
    Dry = A.Dry
    Return
X_Ins_Upd:
    Dim IDry(), UDry(): GoSub X_IDry_UDry
    Set Ins = Drs(Fny, IDry)
    Set Upd = Drs(Fny, UDry)
    
    Return
X_IDry_UDry:
    Dim Dr, IsIns As Boolean, IsUpd As Boolean
    For Each Dr In Dry
        GoSub X_IsIns_IsUpd
        Select Case True
        Case IsIns: Push IDry, Dr
        Case IsUpd: Push UDry, Dr
        End Select
    Next
    Return
X_IsIns_IsUpd:
    IsIns = False
    IsUpd = False
    Dim SkVy(), Sql$, DbDr()
    SkVy = DrSel(Dr, SkIxAy)
    Sql = SelWhSql(T, Fny, Sk, SkVy)
    DbDr = DbqDr(Db, Sql)
    If Sz(DbDr) = 0 Then IsIns = True: Return
    If Not IsEqAy(DbDr, Dr) Then IsUpd = True: Return
    Return
End Sub

Function SelWhSql$(T, Fny$(), Sk$(), SkVy())
SelWhSql = _
"Select " & JnComma(Fny) & vbCrLf & _
"  From [" & T & "]" & _
WhSqp(Sk, SkVy)
End Function
Function DbqDr(A As Database, Q$) As Variant()
DbqDr = RsDr(A.OpenRecordset(Q))
End Function
Function RsDr(A As DAO.Recordset) As Variant()
If Not A.EOF Then RsDr = ItrPrpAy(A.Fields, "Value")
End Function
Sub DrsInsDbt(A As Drs, Db As Database, T)
GoSub X
Dim Sqy$(): GoSub X_Sqy
DbSqyRun Db, Sqy
Exit Sub
X:
    Dim Dry, Fny$(), Sk$()
    Fny = A.Fny
    Dry = A.Dry
    Sk = DbtSk(Db, T)
    Return
X_Sqy:
    Dim Dr
    For Each Dr In AyNz(Dry)
        Push Sqy, InsSql(T, Fny, Dr)
    Next
    Return
End Sub
Function DrSel(A, IxAy) As Variant()
Dim Ix
For Each Ix In IxAy
    Push DrSel, A(Ix)
Next
End Function
Function DrySel(A, IxAy) As Variant()
Dim Dr
For Each Dr In AyNz(A)
    PushI DrySel, DrSel(Dr, IxAy)
Next
End Function
Function DryWhIxAyValAy(A, IxAy, ValAy) As Variant()
Dim Dr
For Each Dr In A
    If IsEqAy(DrSel(Dr, IxAy), ValAy) Then PushI DryWhIxAyValAy, Dr
Next
End Function
Function DryPkMinus(A, B, PkIxAy&()) As Variant()
Dim AK(): AK = DrySel(A, PkIxAy)
Dim BK(): BK = DrySel(B, PkIxAy)
Dim CK(): CK = DryPkMinus(AK, BK, PkIxAy)
DryPkMinus = DryWhIxAyValAy(A, PkIxAy, CK)
End Function
Function DrsPkDiff(A As Drs, B As Drs, PkSs$) As Drs

End Function
Function DrsPkMinus(A As Drs, B As Drs, PkSs$) As Drs
Dim Fny$(), PkIxAy&()
Fny = A.Fny: If Not IsEqAy(Fny, B.Fny) Then Stop
PkIxAy = AyIxAy(Fny, SslSy(PkSs))
Set DrsPkMinus = Drs(Fny, DryPkMinus(A.Dry, B.Dry, PkIxAy))
End Function

Sub DrsUpdDbt(A As Drs, Db As Database, T)
Dim Sqy$(): GoSub X
DbSqyRun Db, Sqy
Exit Sub
X:
    Dim Dr, Fny$(), Dry(), Sk$()
    Fny = A.Fny
    Sk = DbtSk(Db, T)
    Dry = A.Dry
    For Each Dr In AyNz(Dry)
        Push Sqy, UpdSqlFmt(T, Sk, Fny, Dr)
    Next
    Return
End Sub

Private Sub Z_SetSqpFmt()
Dim Fny$(), Vy()
Ept = RplVBar("|  Set|" & _
"    [A xx] = 1                     ,|" & _
"    B      = '2'                   ,|" & _
"    C      = #2018-12-01 12:34:56# ")
Fny = LinTermAy("[A xx] B C"): Vy = Array(1, "2", #12/1/2018 12:34:56 PM#): GoSub Tst
Exit Sub
Tst:
    Act = SetSqp(Fny, Vy)
    C
    Return
End Sub
Function IsSqBktQuoteNeeded(A) As Boolean
If IsSqBktQuoted(A) Then Exit Function
IsSqBktQuoteNeeded = True
If HasSpc(A) Then Exit Function
If HasDot(A) Then Exit Function
If HasPound(A) Then Exit Function
IsSqBktQuoteNeeded = False
End Function
Function HasDot(A) As Boolean
HasDot = InStr(A, ".") > 0
End Function
Function HasPound(A) As Boolean
HasPound = InStr(A, "#") > 0
End Function

Function SqBktQuoteIfNeeded$(A)
If IsSqBktQuoteNeeded(A) Then
    SqBktQuoteIfNeeded = "[" & A & "]"
Else
    SqBktQuoteIfNeeded = A
End If
End Function
Function AySqBktQuoteIfNeeded(A) As String()
Dim X
For Each X In AyNz(A)
    PushI AySqBktQuoteIfNeeded, SqBktQuoteIfNeeded(X)
Next
End Function

Function FnyAlignQuote(Fny$()) As String()
FnyAlignQuote = AyAlignL(AySqBktQuoteIfNeeded(Fny))
End Function

Function SqlValAyQuote(Vy) As String()
Dim V
For Each V In Vy
    PushI SqlValAyQuote, SqlValQuote(V)
Next
End Function
Function SetSqp$(Fny$(), Vy())
Dim A$: GoSub X_A
SetSqp = "  Set " & A
Exit Function
X_A:
    Dim L$(): L = AySqBktQuoteIfNeeded(Fny)
    Dim R$(): R = SqlValAyQuote(Vy)
    Dim J%, O$()
    For J = 0 To UB(L)
        Push O, L(J) & " = " & R(J)
    Next
    A = JnComma(O)
    Return
End Function
Function SetSqpFmt$(Fny$(), Vy())
Dim A$: GoSub X_A
SetSqpFmt = vbCrLf & "  Set" & vbCrLf & A
Exit Function
X_A:
    Dim L$(): L = FnyAlignQuote(Fny)
    Dim R$(): GoSub X_R
    Dim J%, O$(), S$
    S = Space(4)
    For J = 0 To UB(L)
        Push O, S & L(J) & "= " & R(J)
    Next
    A = JnCrLf(O)
    Return
X_R:
    R = AyAlignL(SqlValAyQuote(Vy))
    Dim J1%
    For J1 = 0 To UB(R) - 1
        R(J1) = R(J1) + ","
    Next
    Return
End Function

Function SqlValQuote$(A)
Dim O$
Select Case True
Case IsStr(A): O = "'" & Replace(A, "'", "''") & "'"
Case IsNumeric(A): O = A
Case IsDate(A): O = "#" & Format(A, "YYYY-MM-DD HH:MM:SS") & "#"
Case IsEmpty(A): O = "null"
End Select
SqlValQuote = O
End Function

Function InsSql$(T, Fny$(), Dr)
Dim A$, B$
A = JnComma(Fny)
B = JnComma(AyMapSy(Dr, "SqlValQuote"))
InsSql = FmtQQ("Insert Into [?] (?) Values(?)", T, A, B)
End Function

Function WhSqpFmt$(Fny$(), Vy)
Dim R$(): R = AyAlignL(SqlValAyQuote(Vy))
Dim L$(): L = FnyAlignQuote(Fny)
Dim Ay$(), J%
For J = 0 To UB(L)
    Push Ay, L(J) & "= " & R(J)
Next
For J = 0 To UB(L) - 1
    Ay(J) = Ay(J) & "And"
Next
WhSqpFmt = vbCrLf & "  Where" & vbCrLf & JnCrLf(AyAddPfx(Ay, "    "))
End Function
Function WhSqp$(Fny$(), Vy)
Dim R$(): R = SqlValAyQuote(Vy)
Dim L$(): L = AySqBktQuoteIfNeeded(Fny)
Dim Ay$(), J%
For J = 0 To UB(L)
    Push Ay, L(J) & "= " & R(J)
Next
WhSqp = vbCrLf & "  Where " & Join(Ay, " and ")
End Function

Sub Z_UpdSqlFmt()
Dim T$, Sk$(), Fny$(), Dr
T = "A"
Sk = LinTermAy("X Y")
Fny = LinTermAy("X Y A B C")
Dr = Array(1, 2, 3, 4, 5)
Ept = "Update [A]" & _
vbCrLf & "  Set" & _
vbCrLf & "    A = 3 ," & _
vbCrLf & "    B = 4 ," & _
vbCrLf & "    C = 5 " & _
vbCrLf & "  Where" & _
vbCrLf & "    X = 1 And" & _
vbCrLf & "    Y = 2 "
GoSub Tst

T = "A"
Sk = LinTermAy("[A 1] B CD")
Fny = LinTermAy("X Y B Z CD [A 1]")
Dr = Array(1, 2, 3, 4, "XX", #1/2/2018 12:34:00 PM#)
Ept = "Update [A]" & _
vbCrLf & "  Set" & _
vbCrLf & "    X = 1 ," & _
vbCrLf & "    Y = 2 ," & _
vbCrLf & "    Z = 4 " & _
vbCrLf & "  Where" & _
vbCrLf & "    [A 1] = #2018-01-02 12:34:00# And" & _
vbCrLf & "    B     = 3                     And" & _
vbCrLf & "    CD    = 'XX'                  "
GoSub Tst
Exit Sub
Tst:
    Act = UpdSql(T, Sk, Fny, Dr)
    C
    Return
End Sub

Function UpdSql$(T, Sk$(), Fny$(), Dr)
If Sz(Sk) = 0 Then Stop
Dim Upd$, Set_$, Wh$: GoSub X_Upd_Set_Wh
UpdSql = Upd & Set_ & Wh
Exit Function
X_Upd_Set_Wh:
    Dim Fny1$(), Dr1(), SkVy(): GoSub X_Fny1_Dr1_SkVy
    Upd = "Update [" & T & "]"
    Set_ = SetSqp(Fny1, Dr1)
    Wh = WhSqpFmt(Sk, SkVy)
    Return
X_Ay:
    Dim L$(), R$()
    L = FnyAlignQuote(Fny)
    R = SqlValAyQuote(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, IxAy%(), I%
    For Each Ski In Sk
        I = AyIx(Fny, Ski)
        If I = -1 Then Stop
        Push IxAy, I
        Push SkVy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not AyHas(IxAy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Function UpdSqlFmt$(T, Sk$(), Fny$(), Dr)
If Sz(Sk) = 0 Then Stop
Dim Upd$, Set_$, Wh$: GoSub X_Upd_Set_Wh
UpdSqlFmt = Upd & Set_ & Wh
Exit Function
X_Upd_Set_Wh:
    Dim Fny1$(), Dr1(), SkVy(): GoSub X_Fny1_Dr1_SkVy
    Upd = "Update [" & T & "]"
    Set_ = SetSqp(Fny1, Dr1)
    Wh = WhSqpFmt(Sk, SkVy)
    Return
X_Ay:
    Dim L$(), R$()
    L = FnyAlignQuote(Fny)
    R = SqlValAyQuote(Dr)
    Return
X_Fny1_Dr1_SkVy:
    Dim Ski, J%, IxAy%(), I%
    For Each Ski In Sk
        I = AyIx(Fny, Ski)
        If I = -1 Then Stop
        Push IxAy, I
        Push SkVy, Dr(I)    '<====
    Next
    Dim F
    For Each F In Fny
        If Not AyHas(IxAy, J) Then
            Push Fny1, F        '<===
            Push Dr1, Dr(J)     '<===
        End If
        J = J + 1
    Next
    Return
End Function

Function FbDb(A$) As Database
Set FbDb = DAO.DBEngine.OpenDatabase(A)
End Function

Function DbtSk(A As Database, T) As String()
DbtSk = IdxFny(DbtSkIdx(A, T))
End Function

Function ItrFstNm(A, Nm$)
Dim X
For Each X In A
    If X.Name = Nm Then Set ItrFstNm = X: Exit Function
Next
End Function

Function DbNm$(A As Database)
DbNm = A.Name
End Function
Function DbtSkIdx(A As Database, T) As DAO.Index
Dim O As DAO.Index
Set O = ItrFstNm(A.TableDefs(T).Indexes, "SecondaryKey")
If IsNothing(O) Then Exit Function
If Not O.Unique Then FunEr "DbtSkIdx", "[T] of [Db] has Idx-SecondaryKey.  It should be Unique", DbNm(A), T
If O.Primary Then FunEr "DbtSkIdx", "[T] of [Db] is Primary, but is has a name-SecondaryKey.", DbNm(A), T
Set DbtSkIdx = O
End Function
Sub FunEr(FunNm$, QStr$, ParamArray Ap())

End Sub
Function IdxFny(A As DAO.Index) As String()
If IsNothing(A) Then Exit Function
IdxFny = ItrNy(A.Fields)
End Function

Function LoFny(A As ListObject) As String()
If Not IsNothing(A) Then LoFny = ItrNy(A.ListColumns)
End Function
Function LoDrs(A As ListObject) As Drs
Set LoDrs = Drs("Mth Md", LoDry(A))
End Function
Function LoDry(A As ListObject) As Variant()
LoDry = SqDry(A.DataBodyRange.Value)
End Function
Function LoDrySel(A As ListObject, Fldss$) As Variant() _
' Return as many column as fields in [Fldss] from Lo[A]
Dim IxAy&(), Dry(): GoSub X_IxAy_Dry
Dim Dr
For Each Dr In AyNz(Dry)
    PushI LoDrySel, DrSel(Dr, IxAy)
Next
Exit Function
X_IxAy_Dry:
    Dim Fny$()
    Fny = LoFny(A)
    Dry = LoDry(A)
    IxAy = AyIxAy(Fny, SslSy(Fldss))
    Return
End Function
Function SqRowDr(A, R&) As Variant()
Dim C%
For C = 1 To UBound(A, 2)
    PushI SqRowDr, A(R, C)
Next
End Function
Function SqDry(A) As Variant()
If Not IsArray(A) Then
    SqDry = Array(Array(A))
    Exit Function
End If
Dim R&
For R = 1 To UBound(A, 1)
    PushI SqDry, SqRowDr(A, R)
Next
End Function

Function WbLo(A As Workbook, LoNm$) As ListObject
Dim Ws As Worksheet, O As ListObject
For Each Ws In A.Sheets
    Set O = WsLo(Ws, LoNm)
    If Not IsNothing(O) Then Set WbLo = O: Exit Function
Next
End Function

Function WsLo(A As Worksheet, LoNm$) As ListObject
Dim O As ListObject
For Each O In A.ListObjects
    If O.Name = LoNm Then Set WsLo = O: Exit Function
Next
End Function

Function DrsNRow&(A As Drs)
DrsNRow = Sz(A.Dry)
End Function


Function FbCat(A$) As ADOX.Catalog
Dim O As New Catalog
Set O.ActiveConnection = FbCn(A)
Set FbCat = O
End Function
Function FbCrt(A$) As Database
Set FbCrt = DAO.DBEngine.CreateDatabase(A, dbLangGeneral)
End Function
Sub FbEns(A$)
If Not FfnIsExist(A) Then FbCrt A
End Sub
Function FbAdoCnStr$(A)
'Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
Const C$ = "Provider=Microsoft.ACE.OLEDB.16.0;User ID=Admin;Data Source=?;Mode=Share Deny None;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False"
FbAdoCnStr = FmtQQ(C, A)
End Function

Function FbHasTbl(A$, T$) As Boolean
FbHasTbl = ItrHasNm(FbCat(A).Tables, T)
End Function
Function FbCn(A$) As ADODB.Connection
Dim O As New ADODB.Connection
O.ConnectionString = FbAdoCnStr(A)
O.Open
Set FbCn = O
End Function
Function DrsWhCCNe(A As Drs, C1$, C2$) As Drs
Dim Fny$()
Fny = A.Fny
Set DrsWhCCNe = Drs(Fny, DryWhCCNe(A.Dry, AyIx(Fny, C1), AyIx(Fny, C2)))
End Function

Function DryWhCCNe(A, C1, C2) As Variant()
Dim Dr
For Each Dr In A
    If Dr(C1) <> Dr(C2) Then PushI DryWhCCNe, Dr
Next
End Function

Sub DbSqyRun(A As Database, Sqy$())
Dim Sql
For Each Sql In AyNz(Sqy)
    A.Execute Sql
Next
End Sub

Function MdDclLines$(A As CodeModule)
If A.CountOfDeclarationLines = 0 Then Exit Function
MdDclLines = A.Lines(1, A.CountOfDeclarationLines)
End Function

Function WsSetCdNmAndLoNm(A As Worksheet, Nm$) As Worksheet
Set WsSetCdNmAndLoNm = WsSetLoNm(WsSetCdNm(A, "Ws_" & Nm), Nm)
End Function

Sub Z_WsCmp()
Dim C As VBComponent
Dim Ws As Worksheet
Set Ws = NewWs
Set C = WsCmp(Ws)
Stop

End Sub
Function WsCmp(A As Worksheet) As VBComponent
Dim Wb As Workbook, Pj As VBProject
Set Wb = WsWb(A)
Set Pj = WbPj(Wb)
Set WsCmp = ItrFstNm(Pj.VBComponents, A.CodeName)
End Function

Function WbCmp(A As Workbook) As VBComponent
Set WbCmp = ItrFstNm(WbPj(A).VBComponents, A.CodeName)
End Function

Sub Z_WbPj()
Dim Wb As Workbook, Pj As VBProject
Set Wb = NewWb
Set Pj = WbPj(Wb)
Stop
End Sub

Function WbFx$(A As Workbook)
Dim F$
F = A.FullName
If F = A.Name Then Exit Function
WbFx = F
End Function
Function ItrFstPrpEqV(A, P, V)
Dim X
For Each X In A
    If ObjPrp(X, P) = V Then Set ItrFstPrpEqV = X: Exit Function
Next
End Function
Function WbPj(A As Workbook) As VBProject
Dim Fx$: Fx = WbFx(A)
If Fx <> "" Then Set WbPj = ItrFstPrpEqV(A.Application.Vbe.VBProjects, "FileName", Fx): Exit Function
Dim Ix%: GoSub X_Ix
Dim Pj, I%
For Each Pj In A.Application.Vbe.VBProjects
    If PjFfn(CvPj(Pj)) = "" Then
        If Ix = I Then
            Set WbPj = Pj
            Exit Function
        End If
        I = I + 1
    End If
Next
Stop
Exit Function
X_Ix:
    Dim Wb As Workbook, P
    P = ObjPtr(A)
    For Each Wb In A.Application.Workbooks
        If Wb.Name = Wb.FullName Then
            If P = ObjPtr(Wb) Then Return
            Ix = Ix + 1
        End If
    Next
    Stop
    Return
End Function

Function WbSetCdNm(A As Workbook, CdNm$) As Worksheet
WbCmp(A).Name = CdNm
Set WbSetCdNm = A
End Function

Function WsSetCdNm(A As Worksheet, CdNm$) As Worksheet
WsCmp(A).Name = CdNm
Set WsSetCdNm = A
End Function

Function ItrFst(A)
Dim X
For Each X In A
    Asg X, ItrFst
    Exit Function
Next
End Function

Function WsSetLoNm(A As Worksheet, Nm$) As Worksheet
Dim Lo As ListObject
Set Lo = ItrFst(A.ListObjects)
If Not IsNothing(Lo) Then Lo.Name = "T_" & Nm
Set WsSetLoNm = A
End Function

Function LoWb(A As ListObject) As Workbook
Set LoWb = WsWb(LoWs(A))
End Function
