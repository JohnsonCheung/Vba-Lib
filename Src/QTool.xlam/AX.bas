Attribute VB_Name = "AX"
Option Explicit
Public AcsClass As New Acs
Type WhNm
    Re As RegExp
    ExlAy() As String
End Type
Type WhMth
    Kd() As String
    Mdy() As String
    Nm As WhNm
End Type
Type WhMd
    Ty() As vbext_ComponentType
    Nm As WhNm
End Type
Public Type WhMdMth
    Md As WhMd
    Mth As WhMth
End Type
Type WhPjMth
    Pj As WhNm
    MdMth As WhMdMth
End Type
Type Either
    IsLeft As Boolean
    Left As Variant
    Right As Variant
End Type
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
Property Get Acs() As Access.Application
Set Acs = AcsClass.V
End Property

Private Sub ZZ_DrsKeyCntDic()
Dim Drs As Drs, Dic As Dictionary
Set Drs = VbeMth12Drs(CurVbe)
Set Dic = DrsKeyCntDic(Drs, "Nm")
DicBrw Dic
End Sub







Private Sub ZZ_DrsGpFlat()
Dim Act As Drs, Drs2 As Drs, Drs1 As Drs, N1%, N2%
Set Drs1 = VbeFun12Drs(CurVbe)
N1 = Sz(Drs1.Dry)
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



Function AlignL$(A, W, Optional ErIFmnotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "AlignL"
If ErIFmnotEnoughWdt And DoNotCut Then
    Stop
    'Er CSub, "Both {ErIFmnotEnoughWdt} and {DontCut} cannot be True", ErIFmnotEnoughWdt, DoNotCut
End If
Dim S$: S = VarStr(A)
AlignL = StrAlignL(S, W, ErIFmnotEnoughWdt, DoNotCut)
End Function


























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

Function DupMthFNy_SamMthBdyFunFNy(A$(), Vbe As Vbe) As String()
Dim Gp(): Gp = DupMthFNy_GpAy(A)
Dim O$(), N, Ny
For Each Ny In Gp
    If DupMthFNyGp_IsDup(Ny) Then
        For Each N In Ny
            Push O, N
        Next
    End If
Next
DupMthFNy_SamMthBdyFunFNy = O
End Function

Function DupMthFNyGp_CmpLy(A, Optional OIx% = -1, Optional OSam% = -1, Optional InclSam As Boolean) As String()
'DupMthFNyGp is Variant/String()-of-MthFNm with all mth-nm is same
'MthFNm is MthNm in FNm-fmt
'          Mth is Prp/Sub/Fun in Md-or-Cls
'          FNm-fmt which is 'Nm:Pj.Md'
'DupMthFNm is 2 or more MthFNy with same MthNm
Ass DupMthFNyGp_IsVdt(A)
Dim J%, I%
Dim LinesAy$()
Dim UniqLinesAy$()
    LinesAy = AyMapSy(A, "FunFNm_MthLines")
    UniqLinesAy = AyWhDist(LinesAy)
Dim MthNm$: MthNm = Brk(A(0), ":").S1
Dim Hdr$(): Hdr = DupMthFNyGp_CmpLy__1Hdr(OIx, MthNm, Sz(A))
Dim Sam$(): Sam = DupMthFNyGp_CmpLy__2Sam(InclSam, OSam, A, LinesAy)
Dim Syn$(): Syn = DupMthFNyGp_CmpLy__3Syn(UniqLinesAy, LinesAy, A)
Dim Cmp$(): Cmp = DupMthFNyGp_CmpLy__4Cmp(UniqLinesAy, LinesAy, A)
DupMthFNyGp_CmpLy = AyAddAp(Hdr, Sam, Syn, Cmp)
End Function

Function DupMthFNyGp_IsVdt(A) As Boolean
If Not IsSy(A) Then Exit Function
If Sz(A) <= 1 Then Exit Function
Dim N$: N = Brk(A(0), ":").S1
Dim J%
For J = 1 To UB(A)
    If N <> Brk(A(J), ":").S1 Then Exit Function
Next
DupMthFNyGp_IsVdt = True
End Function

Function EitherL(A) As Either
Asg A, EitherL.Left
EitherL.IsLeft = True
End Function
Function PjHasCmp(A As VBProject, Nm$) As Boolean
PjHasCmp = ItrHasNm(A.VBComponents, Nm)
End Function
Function PjHasCmpWhRe(A As VBProject, Re As RegExp) As Boolean
PjHasCmpWhRe = ItrHasNmWhRe(A.VBComponents, Re)
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
Function CvDic(A) As Dictionary
Set CvDic = A
End Function
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

Function EmpDicAy() As Dictionary()
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
FmtQQ = O
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



Function FunCmpFmt(FunNm, Optional InclSam As Boolean) As String()
'Found all Fun with given name and compare within curVbe if it is same
'Note: Fun is any-Mdy Fun/Sub/Prp-in-Md
Dim O$()
Dim N$(): N = FunFNmAy(FunNm)
DupMthFNy_ShwNotDupMsg N, FunNm
If Sz(N) <= 1 Then Exit Function
FunCmpFmt = DupMthFNyGp_CmpLy(N, InclSam:=InclSam)
End Function

Function FunFNmAy(FunNm) As String()
Stop '
'FunFNmAy = VbeFunFNm(CurVbe, FunPatn:="^" & FunNm & "$", FunExl:="Z__Tst", WhMdy:="Pub")
End Function
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

Function FxWb(A) As Workbook
Set FxWb = Xls.Workbooks.Open(A)
End Function

Function FxaNm_Fxa$(A)
FxaNm_Fxa = CurPjPth & A & ".xlam"
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

Function IsMthLin(A) As Boolean
IsMthLin = LinMthKd(A) <> ""
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






Function RmvMdy$(A)
RmvMdy = LTrim(RmvPfxAyS(A, MdyAy))
End Function

Function LinRmvT1$(A)
Dim O$: O = A
ShfTerm O
LinRmvT1 = O
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






Function PjMthSq(A As VBProject) As Variant()
PjMthSq = MthKy_Sq(PjMthKy(A, True))
End Function

Function CurMdMthNyWh(A As WhMth) As String()
CurMdMthNyWh = MdMthNyWh(CurMd, A)
End Function

Function CurMdMthNy() As String()
CurMdMthNy = MdMthNy(CurMd)
End Function










Function CurMdSrtRptFmt() As String()
CurMdSrtRptFmt = MdSrtRptFmt(CurMd)
End Function
Sub Mov()
CurMthMov "IdeSrt"
End Sub









Function Min(ParamArray A())
Dim O, J&, Av()
Av = A
Min = AyMin(Av)
End Function





Function CurVbeMthMdDNy(MthNm) As String()
CurVbeMthMdDNy = VbeMthMdDNy(CurVbe, MthNm)
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









Function IsMthExist(A As Mth) As Boolean
IsMthExist = MdHasMth(A.Md, A.Nm)
End Function

Function IsPubMth(A As Mth) As Boolean
Dim L$: L = MthLin(A): If L = "" Then Stop
Dim Mdy$: Mdy = TakMdy(L)
If Mdy = "" Or Mdy = "Public" Then IsPubMth = True
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


Function LinMthFNmWh$(A, B As WhMth)
With B
    Dim L$, M$, T$, N$
    L = A
    M = ShfMdy(L): If Not AySel(.Mdy, M) Then Exit Function
    T = ShfMthTy(L): If Not AySel(.Kd, MthKd(T)) Then Exit Function
    N = TakNm(L): If Not IsNmSel(N, .Nm.Re, .Nm.ExlAy) Then Exit Function
End With
LinMthFNmWh = N & ":" & MthShtTy(T) & "." & ShtMdy(M)
End Function

Function LinMthNmWh$(A, B As WhMth)
LinMthNmWh = Brk1(LinMthFNmWh(A, B), ":").S1
End Function

Sub Z_LinMthFNm()
Dim A$
A = "Function LinMthFNm$(A)": Ept = "LinMthFNm:Fun.": GoSub Tst
Exit Sub
Tst:
    Act = LinMthFNm(A)
    C
    Return
End Sub

Function LinMthNm$(A)
LinMthNm = Brk1(LinMthFNm(A), ":").S1
End Function

Function LinMthFNm$(A)
Dim L$, M$, T$, N$
L = A
M = ShfMdy(L)
T = ShfMthTy(L): If T = "" Then Exit Function
N = TakNm(L): If N = "" Then Exit Function
LinMthFNm = N & ":" & MthShtTy(T) & "." & ShtMdy(M)
End Function

Function LinMthKey$(A$, Optional PjNm$, Optional MdNm$, Optional IsWrap As Boolean)
Dim M$ 'Mdy
Dim S$ 'MthShtTy *Sub *Fun *Get *Let *Set
Dim N$ 'Name
Dim B$()
    B = LinMthBrk(A)
    N = B(2)
    If B(2) = "" Then Stop
    M = B(0): If M = "Pub" Then M = ""
    S = B(1)
Dim P% 'Priority
    Select Case True
    Case IsPfx(N, "Init"): P = 1
    Case N = "Z__Tst":    P = 9
    Case N = "ZZ__Tst":   P = 9
    Case IsPfx(N, "Z_"): P = 9
    Case IsPfx(N, "ZZ_"):  P = 8
    Case IsPfx(N, "Z"):    P = 7
    Case Else:             P = 2
    End Select
Dim O$
    Dim Fmt$, NoPjNmMdNm As Boolean
    NoPjNmMdNm = PjNm = "" And MdNm = ""
    Fmt = IIf(NoPjNmMdNm, "?:?|?:?", "?:?|?:?|?:?")
    If Not IsWrap Then Fmt = Replace(Fmt, "|", ":")
    
    If NoPjNmMdNm Then
        O = FmtQQ(Fmt, P, N, S, M)
    Else
        O = FmtQQ(Fmt, PjNm, MdNm, P, N, S, M)
    End If

LinMthKey = O
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

Function NewWs(Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet
Set O = NewWb.Sheets(1)
If O.Name <> WsNm Then O.Name = WsNm
Set NewWs = O
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
Function Pj(PjNm$) As VBProject
Set Pj = CurVbe.VBProjects(PjNm)
End Function
Function SelOy(A, PrpSsl$) As Variant()

End Function


Function WhNm(Patn$, Exl$) As WhNm
With WhNm
Set .Re = Re(Patn)
.ExlAy = SslSy(Exl)
End With
End Function
Function WhMth(WhMdy$, WhKd$, Patn$, Exl$) As WhMth
With WhMth
    .Kd = CvWhMthKd(WhKd)
    .Mdy = CvWhMdy(WhMdy)
    .Nm = WhNm(Patn, Exl)
End With
End Function
Function WhMd(WhCmpTy$, Patn$, Exl$) As WhMd
WhMd.Ty = CvWhCmpTy(WhCmpTy)
WhMd.Nm = WhNm(Patn, Exl)
End Function
Function WhMdMth(Optional MthPatn$, Optional MthExl$, Optional WhMdy$, Optional WhKd$, Optional MdPatn$, Optional MdExl$, Optional WhCmpTy$) As WhMdMth
With WhMdMth
    .Md = WhMd(WhCmpTy, MdPatn, MdExl)
    .Mth = WhMth(WhMdy, WhKd, MthPatn, MthExl)
End With
End Function
Function PjClsAyWh(A As VBProject, B As WhNm) As CodeModule()
Dim M As WhMd
M.Nm = B
M.Ty = CvCmpTyAy("Cls")
PjClsAyWh = PjMdAyWh(A, M)
End Function
Function SelStdMd() As WhMd
SelStdMd = WhMd("Std", "", "")
End Function
Function SelClsMd() As WhMd
SelClsMd = WhMd("Cls", "", "")
End Function
Function ClsCmp() As vbext_ComponentType()
ClsCmp = CvWhCmpTy("Cls")
End Function
Function StdCmp() As vbext_ComponentType()
StdCmp = CvWhCmpTy("Std")
End Function
Function WhEmpNm() As WhNm
End Function
Function PjStdAy(A As VBProject) As CodeModule()
PjStdAy = PjMdAyWh(A, SelStdMd)
End Function
Function PjClsAy(A As VBProject) As CodeModule()
PjClsAy = PjMdAyWh(A, SelClsMd)
End Function

Function PjStdAyWh(A As VBProject, B As WhNm) As CodeModule()
Stop '
'PjStdAyWh = PjMdAyWh(A, WhMd("Std", B.))
End Function



Function PjCmpNy(A As VBProject) As String()
PjCmpNy = ItrNy(A.VBComponents)
End Function
Function PjCmpNyWh(A As VBProject, Optional WhCmpTy$, Optional Patn$, Optional Exl$) As String()
Dim C() As VBComponent
C = OyWhPrpIn(A.VBComponents, "Type", CvWhCmpTy(WhCmpTy))
PjCmpNyWh = AyWhPatnExl(ItrNy(C), Patn, Exl)
End Function
Function PjClsAndMdNy(A As VBProject, Optional Patn$, Optional Exl$) As String()
PjClsAndMdNy = PjCmpNyWh(A, "Cls Mod", Patn, Exl)
End Function

Function PjClsNy(A As VBProject, Optional Patn$, Optional Exl$) As String()
PjClsNy = PjCmpNyWh(A, "Cls", Patn, Exl)
End Function

Function PjCmp(A As VBProject, Nm) As VBComponent
Set PjCmp = A.VBComponents(Nm)
End Function

Function PjDicOfMthKeyzzzMthLines(A As VBProject) As Dictionary
Dim I
Dim O As New Dictionary
For Each I In PjMdAy(A)
    Set O = DicAdd(O, MdMthKeyLinesDic(CvMd(I)))
Next
Set PjDicOfMthKeyzzzMthLines = O
End Function

Function PjDupMth(A As VBProject, Optional IsSamMthBdyOnly As Boolean) As Drs
Dim N$(): N = PjMthFNm(A)
Stop '
Dim N1$(): 'N1 = MthNyWhDup(N)
If IsSamMthBdyOnly Then
'    N1 = DupMthFNy_SamMthBdyFunFNy(N1, A)
End If
'Set PjDupMth = N1
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

Function PjMthFNm(A As VBProject, Optional MdRe As RegExp, Optional MthRe As RegExp, Optional MthExl$, Optional WhMdyAy, Optional WhKdAy) As String()
Dim O$(), I
For Each I In AyNz(PjMdAy(A)) ', Patn:=MdPatn))
'   PushAy O, MdMthFNm(CvMd(I), MthPatn:=FunRe, ExlAy:=FunExl, WhMdy:=WhMdyA, WhKd:=WhKdAy)
Next
PjMthFNm = O
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
Function ItrWhNm(A, Re As RegExp, ExlAy)
ItrWhNm = ItrWhNmInto(A, Re, ExlAy, EmpAy)
End Function
Function ItrWhWhNm(A, B As WhNm)
ItrWhWhNm = ItrWhNmInto(A, B.Re, B.ExlAy, EmpAy)
End Function
Function ItrWhNmInto(A, Re As RegExp, ExlAy, OInto)
Dim X
ItrWhNmInto = OInto
Erase ItrWhNmInto
For Each X In A
    If IsNmSel(X.Name, Re, ExlAy) Then PushObj ItrWhNmInto, X
Next
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

Function PjMdAyWh(A As VBProject, B As WhMd) As CodeModule()
Dim C
For Each C In AyNz(ItrWhWhNm(A.VBComponents, B.Nm))
    With CvCmp(C)
    If AySel(B.Ty, .Type) Then
        PushObj PjMdAyWh, .CodeModule
    End If
    End With
Next
End Function

Function PjMdAy(A As VBProject) As CodeModule()
PjMdAy = ItrPrpAyInto(A.VBComponents, "CodeModule", PjMdAy)
End Function

Function ItrWhNyInto(A, InNy$(), OInto)
Dim O, X
O = OInto
Erase O
For Each X In A
    If AyHas(InNy, X.Name) Then PushObj X, O
Next
ItrWhNyInto = O
End Function

Function CvWhCmpTy(WhCmpTy$) As vbext_ComponentType()
Dim O() As vbext_ComponentType, I
For Each I In AyNz(SslSy(WhCmpTy))
    Push O, CmpShtToTy(I)
Next
CvWhCmpTy = O
End Function


Function PjCmpAyWh(A As VBProject, Optional Re As RegExp, Optional ExlAy, Optional WhCmpTyAy) As VBComponent()
PjCmpAyWh = ItrWhNmInto(A.VBComponents, Re, ExlAy, PjCmpAyWh)
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

Function PjMdNy(A As VBProject, Optional Re As RegExp, Optional ExlAy, Optional WhCmpTyAy) As String()
Stop '
'PjMdNy = PjCmpNyWh(A, Re, ExlAy)
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

Function PjStdClsNy(A As VBProject) As String()
PjStdClsNy = PjMdNy(A, WhCmpTyAy:=CvWhCmpTy("Std Cls"))
End Function

Function PjMthAyWh(A As VBProject, B As WhMdMth) As Mth()
Dim M
For Each M In AyNz(PjMdAyWh(A, B.Md))
    PushObjAy PjMthAyWh, MdMthAyWh(CvMd(M), B.Mth)
Next
End Function

Function PjMthAy(A As VBProject) As Mth()
Dim M
For Each M In PjMdAy(A)
    PushObjAy PjMthAy, MdMthAy(CvMd(M))
Next
End Function


Function PjMthKy(A As VBProject, Optional IsWrap As Boolean) As String()
PjMthKy = AyMapPXSy(PjMdAy(A), "MdMthKy", IsWrap)
End Function

Function PjMthKySq(A As VBProject) As Variant()
PjMthKySq = MthKy_Sq(PjMthKy(A, True))
End Function

Function CurPjMthNyWh(A As WhPjMth) As String()
Stop
'CurPjMthNy = PjMthNy(CurPj, CvPatn(MthPatn), MthExl, WhMdyAy, WhKdAy, MdPatn, MdExl, WhCmpTy)
End Function

Function PjMthNy(A As VBProject) As String()
Dim I, O$()
For Each I In AyNz(PjMdAy(A))
    PushAy O, MdMthNy(CvMd(I))
Next
O = AyAddPfx(O, A.Name & ".")
PjMthNy = O
End Function

Function PjMthNyWh(A As VBProject, B As WhMdMth) As String()
Dim I, O$()
For Each I In AyNz(PjMdAyWh(A, B.Md))
    PushAy O, MdMthNyWh(CvMd(I), B.Mth)
Next
O = AyAddPfx(O, A.Name & ".")
PjMthNyWh = O
End Function

Function PjMthFNyWh(A As VBProject, B As WhMdMth) As String()
Dim I, O$()
For Each I In AyNz(PjMdAyWh(A, B.Md))
    PushAy O, MdMthFNyWh(CvMd(I), B.Mth)
Next
O = AyAddPfx(O, A.Name & ".")
PjMthFNyWh = O
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
   PushAy O, S1S2_Ly(A(I), W1, W2)
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

Function S1S2_Ly(A As S1S2, W1%, W2%) As String()
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
S1S2_Ly = O
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

Function FTIxFC(A As FTIx) As FmCnt
With A
    Set FTIxFC = FmCnt(.Fmix + 1, .Toix - .Fmix + 1)
End With
End Function

Function FTIxAyFC(A() As FTIx) As FmCnt()
FTIxAyFC = AyMapInto(A, "FTIxFC", FTIxAyFC)
End Function




Function ApLines$(ParamArray Ap())
Dim Av(): Av = Ap
ApLines = Join(AyRmvEmp(Av), vbCrLf)
End Function






Function IsRmkLin(A) As Boolean
IsRmkLin = FstChr(LTrim(A)) = "'"
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

Function CurVbePj(A$) As VBProject
Set CurVbePj = CurVbe.VBProjects(A)
End Function

Function CurVbeFunNm() As String()

End Function

Function PjCmpAy(A As VBProject) As VBComponent()
PjCmpAy = ItrAyInto(A.VBComponents, PjCmpAy)
End Function

Sub CurVbePjMdFmtBrw()
Brw VbePjMdFmt(CurVbe)
End Sub









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

Function CurVbeMthNy() As String()
CurVbeMthNy = VbeMthNy(CurVbe)
End Function
Function CurVbeMthNyWh(A As WhPjMth) As String()
CurVbeMthNyWh = VbeMthNyWh(CurVbe, A)
End Function

Function VblLines$(A)
VblLines = Replace(A, "|", vbCrLf)
End Function

Function WdtAy_HdrLin$(A%())
Dim O$(), W
For Each W In A
    Push O, StrDup("-", W + 2)
Next
WdtAy_HdrLin = "|" + Join(O, "|") + "|"
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

Sub FunCmp(FunNm$, Optional InclSam As Boolean)
D FunCmpFmt(FunNm, InclSam)
End Sub

Sub FunSync(A As Mth, Optional ShwCmpLyAft As Boolean)
Dim Lines$: Lines = MthLines(A)
If Lines = "" Then
    Debug.Print FmtQQ("Give Mth(?) not exist", MthDNm(A))
    Exit Sub
End If
Dim M() As Mth
    M = FunSync__1(A, Lines) ' Mth to be replaced
If Sz(M) = 0 Then Exit Sub
Dim I
For Each I In M
    MthRpl CvMth(I), Lines
Next
If ShwCmpLyAft Then
    FunCmp A.Nm, True
End If
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
Function StrIsEq(A, B) As Boolean
StrIsEq = StrComp(A, B, vbTextCompare) = 0
End Function
Sub ItrDoSub(A, SubNm$)
Dim I
For Each I In A
    CallByName A, SubNm, VbMethod
Next
End Sub






















Function JnSpc$(A)
JnSpc = Join(A, " ")
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



Private Function CNmProperMdNm$(A$)
'Given a [Mth}, return the MdNm which the Mth should be copied to
Stop '
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

Sub SelRg_SetXorEmpty(A As Range)
Dim I
For Each I In A
    
Next
End Sub

Private Property Get ZZFmMthDic() As Dictionary
Set ZZFmMthDic = PjMthDic(ZZFmPj)
End Property

Private Property Get ZZToMthDic() As Dictionary
Set ZZToMthDic = PjMthDic(ZZToPj)
End Property

Private Property Get ZZToPj() As VBProject
Set ZZToPj = Pj("QVb")
End Property

Private Property Get ZZFmPj() As VBProject
Set ZZFmPj = Pj("QTool")
End Property

Function IsFmCntInOrd(A() As FmCnt) As Boolean
Dim J%
For J = 0 To UB(A) - 1
    With A(J)
      If .FmLno + .Cnt > A(J + 1).FmLno Then Exit Function
    End With
Next
IsFmCntInOrd = True
End Function



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


Sub ZZ_PjCompile()
PjCompile CurPj
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

Private Sub ZZ_VbeDmpIsSaved()
VbeDmpIsSaved CurVbe
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
If Not AyHas(O, M) Then Push O, M
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


Sub S1S2AyBrw(A() As S1S2)
Brw S1S2AyFmt(A)
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

Private Function FunSync__1(A As Mth, Lines$) As Mth()
Dim Ny$(): Ny = FunFNmAy(A.Nm)
Dim Ny1$(): Ny1 = AyRmvEle(Ny, MthFNm(A))
If Sz(Ny) <> Sz(Ny1) + 1 Then Stop
Dim O() As Mth, J%, M As Mth, L$
For J = 0 To UB(Ny1)
    Set M = MthFNm_Mth(Ny1(J))
    L = MthLines(M): If L = "" Then Stop
    If L <> Lines Then
        PushObj O, M
    End If
Next
If Sz(O) = 0 Then
    Debug.Print FmtQQ("FunSync: There are ?-Fun(?). All have same lines", Sz(Ny), MthDNm(A))
End If
FunSync__1 = O
End Function

Private Property Get ZZSrc() As String()
ZZSrc = MdSrc(CurMd)
End Property



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




Private Sub ZZ_LinesAyFmt()
Dim A$()
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdsksdf|skldfjdf|sdf")
Push A, RplVBar("ksdjlfdf|sdklfjdsfdf|skldfjdf|lskdf|slkdjf|sdlf||")
Push A, RplVBar("ksdjlfdf|sdklfjsdfdsfdsf|skldsdffjdf")
D LinesAyFmt(A)
End Sub


Private Sub ZZ_LinMthKey()
Dim Ay1$(): Ay1 = SrcMthLinAy(CurSrc)
Dim Ay2$(): Ay2 = AyMapSy(Ay1, "LinMthKey")
S1S2AyBrw AyabS1S2Ay(Ay2, Ay1)
End Sub

Private Sub ZZ_LinMthKey_1()
Const A$ = "Function YYA()"
Debug.Print LinMthKey(A)
End Sub

Private Sub ZZ_FunCmp()
FunCmp "FfnDlt"
End Sub

Private Sub ZZ_SrcDclLinCnt()
Dim B1$(): B1 = CurSrc
Dim B2$(): B2 = SrcSrtedLy(B1)
Dim A1%: A1 = SrcDclLinCnt(B1)
Dim A2%: A2 = SrcDclLinCnt(SrcSrtedLy(B1))
End Sub

Private Sub ZZ_SrcSrtRptFmt()
Brw SrcSrtRptFmt(CurSrc, "Pj", "Md")
End Sub

Private Sub ZZ_SrcSrtedBdyLines()
StrBrw SrcSrtedBdyLines(CurSrc)
End Sub
Function WhEmpPjMth() As WhPjMth
End Function
Private Sub ZZ_VbeDupMthCmpLy()
Brw VbeDupMthCmpLy(CurVbe, WhEmpPjMth)
End Sub

Private Sub ZZ_VbeMthNyWh()
Brw VbeMthNyWh(CurVbe, WhEmpPjMth)
End Sub
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

