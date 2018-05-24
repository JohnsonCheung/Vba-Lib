VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Enum eTstOpt
    eAllValidate = 0
    eValidateAsFldVal = 1
    eValidateAsBetNum = 2
    eValidateAsNum = 3
    eValidateAsFny = 4
End Enum

Private B_Items As New VBA.Collection
Private B_T1$
Private B_IsVF As Boolean
Private B_IsInited As Boolean
Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"
Const M_Dup$ = _
                      "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"

Sub AddLBC(Lx%, B$, C$)
Dim O As New LABC
With O
    .Lx = Lx
    .B = B
    .C = C
    Set .Par = Me
End With
B_Items.Add O
End Sub

Property Get ValidateAsFny() As FnyRslt
Dim A As New LABCsRslt
    A.Init Me
Set A = VdtDupFld(A)
Dim O As New FnyRslt
Set ValidateAsFny = O.Init(UniqFny, A.Er)
End Property

Function Init(ABCAy() As ABC, Optional IsVF As Boolean) As LABCs
If B_IsInited Then PrmEr
B_IsInited = True ' Cannot init once
If AyIsEmp(ABCAy) Then PrmEr
If Not AyIsAllEq(Oy.PrpSy("A")) Then PrmEr
Dim ABC As ABC, I, Lx%
For Each I In ABCAy
    Set ABC = I
    With ABC
        AddLBC Lx, .B, .C
    End With
    Lx = Lx + 1
Next
Set Init = Me
End Function

Function InitByABCLines(ABCLines$, Optional IsVF As Boolean) As LABCs
If ABCLines = "" Then PrmEr
Dim Ay() As ABC, Lin
For Each Lin In SplitLines(ABCLines)
    PushObj Ay, ABC(Lin)
Next
Set InitByABCLines = Init(Ay, IsVF)
End Function

Function InitByT1(T1$, Optional IsVF As Boolean) As LABCs
If B_IsInited Then PrmEr
B_IsInited = True ' Cannot init once
B_T1 = T1
B_IsVF = IsVF
Set InitByT1 = Me
End Function

Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property

Property Get IsVF() As Boolean
IsVF = B_IsVF
End Property

Property Get ItemAy() As LABC
Dim O() As LABC, I
For Each I In B_Items
    PushObj O, I
Next
ItemAy = O
End Property

Property Get Items() As VBA.Collection
Set Items = B_Items
End Property

Property Get N&()
N = B_Items.Count
End Property

Property Get NmRslt() As NmRslt
Dim Nm$, Er$(), Ok$()
Dim O As NmRslt
With O
    .Er = Er
    .Ok = Ok
    .Nm = Nm
End With
End Property

Property Get T1() As String
T1 = B_T1
End Property

Property Get ToStr$()
Dim O$(), I, M As LABC
Push O, "LABCs("
Push O, FmtQQ("IsVF(?)", B_IsVF)
Push O, "Items("
For Each I In Items
    Set M = I
    Push O, M.ToStr
Next
Push O, "Items)"
Push O, "LABCs)"
ToStr = JnCrLf(O)
End Property

Sub TstValidateAsBetNum()
Dim ABCLines$
Dim FnyStr$
Dim FmNum&
Dim ToNum&
    ABCLines = RplVBar("Wdt 10 A B C|Wdt 20 X Y Z")
AyDmp ValidateAsBetNumIO(ABCLines, FnyStr, FmNum, ToNum)
End Sub

Sub TstValidateAsFny()
Dim IsVF As Boolean, ABCLy$(), Fny$()
Dim ABCLines$
Push ABCLy, "Wdt 10 A B C"
Push ABCLy, "Wdt 20 A B D"
Fny = LvsSy("A B C D E")
Dim FmNum&, ToNum&
AyDmp ValidateFnyIO(ABCLines)
End Sub

Property Get UniqFny() As String()
Dim I, M As LABC, O$()
For Each I In B_Items
    Set M = I
    PushNoDupAy O, M.Fny
Next
UniqFny = O
End Property

Function ValidateAsBetNum(Fny$(), FmNum&, ToNum&) As LABCsRslt
Dim OEr$(), IsVF As Boolean 'Always a IsVF=True line
IsVF = True
Dim A As LABCsRslt: Set A = ValidateAsFldVal(Fny)
'Dim A1 As LABCsRslt: Set A1 = ValateBetNum1(A, FmNum, ToNum)
'Dim A4() As LCFV: A4 = ZLABCAy_LCFVAy(A2, Fny, IsVF)
'Dim A5$():        A5 = ZLABCAy_Ly(A3)
'With LABCAy_LCFVRslt_OfBetNum
'    .Ok = A5
''    .LCFVAy = A4
'    .Er = OEr
'End With
End Function

Function ValidateAsBetNumIO(ABCLines$, FnyStr$, FmNum&, ToNum&) As String()
Stop
Dim A As New LABCs
A.InitByABCLines ABCLines, IsVF:=True
Dim R As LABCsRslt
Dim O$()
Dim Fny$(): Fny = LvsSy(FnyStr)
Set R = A.ValidateAsBetNum(Fny, FmNum, ToNum) '<========================
   
PushAp O, "LABCAy_LCFVRsltOfBetNum======================"
PushAp O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", ToStr
PushAp O, "Inp2::Fny <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", FnyStr
PushAp O, "Inp3::FmNum ToNum <<<<<<<<<<<<<<<<<<<<<<<<<<<", FmtQQ("FmToNum(? ?)", FmNum, ToNum)
PushAp O, "Oup1::Ok >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.ToStr
PushAp O, "LABCAy_LCFVRsltOfBetNum======================"
PushAp O, ""
ValidateAsBetNumIO = O
End Function

Function ValidateAsFldVal(Fny$()) As LABCsRslt
Dim A1 As LABCsRslt: Set A1 = VdtErFld(Fny)
Dim A2 As LABCsRslt: Set A2 = VdtDupFld(A1)
Set ValidateAsFldVal = A2
End Function

Function ValidateAsFldValIO(ABCLines$, IsVF As Boolean, FnyStr$) As String()
Dim LABCAy() As LABC
Dim O$(), Fny$()
Fny = LvsSy(FnyStr)
InitByABCLines ABCLines, IsVF
PushAp O, "LABCs.ValidateAsFldVal '(===================="
PushAp O, "LABCs.ToStr <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", ToStr
PushAp O, "Oup1::ValidateAsFldVal >>>>>>>>>>>>>>>>>>>>>>", ValidateAsFldVal(Fny).ToStr
PushAp O, "LABCs.ValidateAsFld ')======================="
PushAp O, ""
ValidateAsFldValIO = O
End Function

Function ValidateNmIO(ABCLines$)
Dim LABCAy() As LABC
Dim R1 As NmRslt
Dim O$()
Dim I1$(), O1$(), O2$(), O3$()
'    LABCAy = ZABCLy_LABCAy(ABCLy)
'    R1 = LABCAy_NmRslt(LABCAy) '<========================
    
    I1 = ZLABCAy_Ly(LABCAy)
    O1 = ApSy(R1.Nm)
    O2 = R1.Ok
    O3 = R1.Er
         
         Push O, "LABCAy_NmRslt ==============================="
    PushItmAy O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", I1
    PushItmAy O, "Oup1::Nm >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O1
    PushItmAy O, "Oup2::LABCAy Nm >>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O2
    PushItmAy O, "Oup3::Er >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", O3
         Push O, "LABCAy_NmRslt ==============================="
         Push O, ""
ValidateNmIO = O
End Function

Friend Sub TstValidateAsFldVal()
Dim IsVF As Boolean, ABCLines$, FnyStr$
IsVF = True
ABCLines = _
    "Wdt 10 A B C X" & vbCrLf & _
    "Wdt 20 A B D Y A"
FnyStr = "A B C D E X"
Debug.Print "LABCs.TstValidateAsFldVal"
AyDmp ValidateAsFldValIO(ABCLines, IsVF, FnyStr)
End Sub

Friend Sub TstValidateAsNm()
Dim InpStr$
AyDmp ValidateNmIO(InpStr)
End Sub

Friend Function ValidateFnyIO(ABCLines$) As String()
InitByABCLines ABCLines, IsVF:=True
Dim R As FnyRslt: Set R = ValidateAsFny
Dim O$()
PushAp O, "LABCAy_FnyRslt =============================="
PushAp O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", ToStr
PushAp O, "Oup1::Fny >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", JnSpc(R.Fny)
PushAp O, "Oup2::Lin$ >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.Lin
PushAp O, "Oup3::Er >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.Er.ToStr
PushAp O, "LABCAy_FnyRslt =============================="
PushAp O, ""
ValidateFnyIO = O
End Function

Private Property Get Oy() As Oy
Dim O As New Oy
Set Oy = O.InitByColl(B_Items)
End Property

Private Function ValidateAsBetNum1(A As LABCsRslt, FmNum&, ToNum&, OEr$()) As LABC()
'Assume LABC is always are IsVF=true

End Function

Private Function VdtDupFld1(A() As LABC, J%) As StrRslt
'A(J).C is FldLvs
'Check If it has any duplicated element is same line or following lines
'if yes, remove it by assigning the no-dup-Fny1 to OFldLvs
'        and report the duplication in OEr
Dim Fny$()
    Dim FldLvs$
    FldLvs = A(J).C
    Fny = LvsSy(FldLvs) '<--- All Fld to be checked
Dim O As New StrRslt
Dim OFny$()             '<-- Those Fld has no duplicated will be put in the string result
    Dim I%
    Dim F$
    For I = 0 To UB(Fny)
        F = Fny(I)
        Dim IsDup As Boolean
        Dim DupAtLx% ' Duplicated at which line
            IsDup = VdtDupFld1IsDup(Fny, I)
            If IsDup Then
                DupAtLx = J 'At same line
            Else
                IsDup = VdtDupFld2(A, J, F, _
                    DupAtLx) 'DupAtLx at following line of J+1...
            End If
        
        If IsDup Then
            Dim Lx%, Msg$
            Lx = A(J).Lx
            Msg = FmtQQ(M_Fld_IsDup, Lx, F, DupAtLx)
            O.Er.AddMsg Msg    '<== Report duplication in OEr
        Else
            Push OFny, F    '<== Push to Fny1 for no-dup
        End If
    Next
Set VdtDupFld1 = O
End Function

Private Function VdtDupFld1IsDup(Fny$(), I%) As Boolean
'Check if Fny(I)-element has duplication found in Fny(I+1..)
Dim F$: F = Fny(I)
Dim II%
For II = I + 1 To UB(Fny)
    If Fny(II) = F Then VdtDupFld1IsDup = True: Exit Function
Next
End Function

Private Function VdtDupFld2(A() As LABC, J%, F$, _
    ODupAtLx%) As Boolean
'Check if F has duplicated-element found in A(J+1...)
Dim JJ%, Fny$(), FldLvs$
ODupAtLx = -1
'For JJ = J + 1 To LABC_UB(A)
'    FldLvs = A(JJ).C
'    Fny = LvsSy(FldLvs)
'    If AyHas(Fny, F) Then
'        ODupAtLx = JJ
'        VdtDupFld2 = True
'        Exit Function
'    End If
'Next
End Function

Private Function VdtDupFld(A As LABCsRslt) As LABCsRslt
Dim A1 As LABCs
Dim A2() As LABC
Dim LABCs As New LABCs
    Set A1 = A.LABCs
    If A1.IsEmp Then Set VdtDupFld = A: Exit Function
    LABCs.InitByT1 A1.T1, A1.IsVF
    A2 = A1.ItemAy

Dim J%, I
Dim Ok As StrRslt  ' Those field is not duplicated, will be in Ok as a StrRslt.
                   ' Any field has duplication, will be put in Ok.Er
Dim O As New LABCsRslt
If A1.IsVF Then
    For J = 0 To UB(A2)
        Set Ok = VdtDupFld1(A2, J)
        If Ok.Str <> "" Then
            With A2(J)
                LABCs.AddLBC .Lx, .B, Ok.Str
            End With
        End If
        O.Er.Add Ok.Er
    Next
Else
    For J = 0 To UB(A2)
        Set Ok = VdtDupFld3(A2, J)
        If Ok.Str <> "" Then
            With A2(J)
                LABCs.AddLBC .Lx, Ok.Str, .B
            End With
        End If
        O.Er.Add Ok.Er
    Next
End If

Set O.LABCs = LABCs
Set VdtDupFld = O
End Function

Private Function VdtDupFld3(A() As LABC, J%) As StrRslt
Stop
End Function

Private Sub VdtErFld1(Lx%, FldLvs$, Fny$(), _
    OFldLvs$, OEr$())
Dim Fny1$(), Fny2$(), J%, F$, Msg$
Fny1 = LvsSy(FldLvs)
For J = 0 To UB(Fny1)
    F = Fny1(J)
    If Not AyHas(Fny, F) Then
        Msg = FmtQQ(M_Fld_IsInValid, Lx, F)
        Push OEr, Msg   '<=================
    Else
        Push Fny2, F
    End If
Next
OFldLvs = JnSpc(Fny2)
End Sub

Private Function VdtErFld(Fny$()) As LABCsRslt
Dim J%, LABCs As New LABCs, Er As New Er, FldLvs$, M As LABC, Lx%, C$, F$
If IsVF Then
'    For J = 0 To LABC_UB(A)
'        With A(J)
'            Lx = .Lx
'            C = .C
'        End With
'
'        VdtErFld1 Lx, C, Fny, _
'            FldLvs, OEr
'
'        If FldLvs <> "" Then
'            M = A(J)
'            M.C = FldLvs
'            LABC_Push O, M  '<============
'        End If
'    Next
'    ZErOf_InvalidFld = O
    Exit Function
End If
Dim Msg$
'For J = 0 To LABC_UB(A)
'    With A(J)
'        Lx = .Lx
'        F = .B
'    End With
'    If AyHas(Fny, F) Then
'        LABC_Push O, A(J)       '<======
'    Else
'        Msg = FmtQQ(M_Fld_IsInValid, Lx, F)
'        Push OEr, Msg           '<======
'    End If
'Next
Dim O As New LABCsRslt
With O
    Set .LABCs = LABCs
    Set .Er = Er
End With
Set VdtErFld = O
End Function

Private Function ZLABCAy_Ly(A() As LABC) As String()
Dim J%, O$()
'For J = 0 To LABC_UB(A)
'    Push O, ZLABC_Lin(A(J))
'Next
ZLABCAy_Ly = O
End Function

Private Function ZLABC_Lin$(A As LABC)
With A
'    ZLABC_Lin = FmtQQ("? ? ? ?", .Lx, .A, .B, .C)
End With
End Function

Private Function ZLCFVAy_LABCAy(A() As LCFV, T1$, IsVF As Boolean) As LABC()
Dim O() As LABC, M As LABC
If IsVF Then
    ZLCFVAy_LABCAy = ZLCFVAy_LABCAy_1(A, T1)
    Exit Function
End If
End Function

Private Function ZLCFVAy_LABCAy_1(A() As LCFV, T1$) As LABC()
Dim J%, Lx%, M As LABC, LxAy%(), V$, FldLvs$, O() As LABC
'LxAy = ZLxAy(A)
For J = 0 To UB(LxAy)
Next
With M
'    .A = T1
    .B = V
    .C = FldLvs
    .Lx = Lx
End With
'LABC_Push O, M
End Function

Private Function ZLCFVAy_Ly(A() As LCFV) As String()
Dim J%, O$()
'For J = 0 To LCFV_UB(A)
'    Push O, ZLCFV_Lin(A(J))
'Next
ZLCFVAy_Ly = O
End Function

Private Function ZLCFV_Lin$(A As LCFV)
With A
    ZLCFV_Lin = FmtQQ("? ? ? ?", .Lx, .Cno, .F, .V)
End With
End Function

Private Function ZLinLABC(Lin$, Lx%) As LABC
Dim A$, B$, C$
LinAsgTTRst Lin, A, B, C
With ZLinLABC
    .Lx = Lx
'    .A = A
    .B = B
    .C = C
End With
End Function

Friend Sub Tst(Optional Opt As eTstOpt = eValidateAsFldVal)
Select Case Opt
Case eValidateAsFldVal: TstValidateAsFldVal
Case eValidateAsBetNum: TstValidateAsBetNum
Case eValidateAsNm:     TstValidateAsNm
Case eValidateAsFny:    TstValidateAsFny
Case eAllValidate:
    TstValidateAsFldVal
    TstValidateAsBetNum
    TstValidateAsNm
    TstValidateAsFny
Case Else
    PrmEr
End Select
End Sub
