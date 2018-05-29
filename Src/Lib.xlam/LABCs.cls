VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Enum eTstOpt
    eAllValidate = 0
    eValidateAsFldVal = 1
    eValidateAsBetNum = 2
    eValidateAsNum = 3
    eValidateAsFny = 4
End Enum

Private B_Ay() As LABC
Private B_T1$
Private B_IsVF As Boolean
Private B_IsInited As Boolean
Const M_Should_Lng$ = "Lx(?) Fld(?) should have val(?) be a long number"
Const M_Should_Num$ = "Lx(?) Fld(?) should have val(?) be a number"
Const M_Should_Bet$ = "Lx(?) Fld(?) should have val(?) be between (?) and (?)"
Const M_Dup$ = _
                      "Lx(?) Fld(?) is found duplicated in Lx(?).  This item is ignored"
Friend Property Get CnoAyWdtAy(Fny$()) As CnoAyWdtAy
Dim O As New CnoAyWdtAy
With CnoAyValAy(Fny)
O.SetCnoAyWdtAy .CnoAy, AyIntAy(.ValAy)
End With
Set CnoAyWdtAy = O
End Property
Private Property Get CnoAyValAy(Fny$()) As CnoAyValAy
Dim OCno%(), OVal(), V
Dim I%, Fny1$(), Cno%
Dim J%, JJ%
For J = 0 To UB(B_Ay)
    I = B_Ay(J).C
    V = B_Ay(J).B
    Fny1 = B_Ay(J).Fny
    For JJ = 0 To UB(Fny1)
        Cno = AyIx(Fny, Fny1(JJ))
        Push OCno, I
        Push OVal, V
    Next
Next
Dim O As New CnoAyValAy
O.SetCnoAyValAy OCno, OVal
End Property
Sub AddLBC(Lx%, B$, C$)
Dim O As New LABC
With O
    .Lx = Lx
    .B = B
    .C = C
    Set .Par = Me
End With
Push B_Ay, O
End Sub

Property Get ValidateAsFny() As FnyRslt
Dim A As New LABCsRslt
Set A = VdtDupFld(A)
Dim O As New FnyRslt
Set ValidateAsFny = O.Init(UniqFny, A.Er)
End Property

Function Init(ABCAy() As ABC, Optional IsVF As Boolean) As LABCs
If B_IsInited Then PrmEr
B_IsInited = True ' Cannot init once
If AyIsEmp(ABCAy) Then PrmEr
If Not AyIsAllEq(Oy.PrpSy("A")) Then PrmEr
B_T1 = ABCAy(0).A
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

Function InitByLines(ABCLines$, Optional IsVF As Boolean) As LABCs
If ABCLines = "" Then PrmEr
Dim Ay() As ABC, Lin
For Each Lin In SplitLines(ABCLines)
    PushObj Ay, ABC(Lin)
Next
Set InitByLines = Init(Ay, IsVF)
End Function

Function InitByT1(T1$, Optional IsVF As Boolean, Optional ABLy0) As LABCs
Dim ABLy$(): ABLy = DftNy(ABLy0)
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

Property Get Ay() As LABC()
Ay = B_Ay
End Property

Property Get N&()
N = Sz(B_Ay)
End Property

Private Property Get Nm$()
If IsEmp Then Nm = "?": Exit Function
Dim A As LABC: Set A = B_Ay(0)
Dim T1$: T1 = Lin(A.C).T1
If T1 = "" Then Nm = "?": Exit Function
Nm = T1
End Property

Property Get NmRslt() As NmRslt
Dim O As New NmRslt
With O
    .Er.AddErLyAp NmErNoLin, NmErExcessLin, NmErMultiName
    .Nm = Nm
End With
End Property

Private Function NmErNoLin() As String()
If IsEmp Then NmErNoLin = ApSy(FmtQQ("There is not ?-line", C2_Lo_Nam))
End Function

Property Get U&()
U = N - 1
End Property

Private Function NmErExcessLin() As String()
Dim J%, O$()
If N <= 1 Then Exit Function
For J = 1 To U
    Push O, FmtQQ(M_Nm_ExcessLin, B_Ay(J).Lx)
Next
NmErExcessLin = O
End Function

Private Function NmErMultiName() As String()
If IsEmp Then Exit Function
Dim A As LABC: Set A = B_Ay(0)

End Function
Property Get T1() As String
T1 = B_T1
End Property
Property Get Ly() As String()

End Property
Property Get ToStr$()
Dim O$(), I, M As LABC
Push O, "LABCs("
Push O, FmtQQ("IsVF(?)", B_IsVF)
If IsEmp Then
    Push O, "Items()"
Else
    Push O, "Items("
    For Each I In B_Ay
        Set M = I
        Push O, M.ToStr
    Next
    Push O, "Items)"
End If
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

Property Get UniqFny() As String()
Dim I, M As LABC, O$()
If IsEmp Then Exit Property
For Each I In B_Ay
    Set M = I
    PushNoDupAy O, M.Fny
Next
UniqFny = O
End Property

Function ValidateAsBetNum(Fny$(), FmNum&, ToNum&) As LABCsRslt
Dim OEr$(), IsVF As Boolean 'Always a IsVF=True line
IsVF = True
Dim A As LABCsRslt: Set A = ValidateAsFldVal(Fny)
Dim A1() As LABC:       A1 = A.LABCs.Ay   ' To be validated
Dim A2 As New LABCs:         A2.InitByT1 A.LABCs.T1, A.LABCs.IsVF
Dim Er As New Er
    Dim J%, V&, B$, C$, Lx%, Msg$
    For J = 0 To UB(A1)
        With A1(J)
            Lx = .Lx
            B = .B
            C = .C
        End With
        If Not IsNum(B) Then
            Msg = FmtQQ(M_Val_IsNonNum, Lx, B)
            Er.AddMsg Msg
            GoTo Nxt
        End If
        V = B
        If FmNum > V Or V > ToNum Then
            Msg = FmtQQ(M_Val_ShouldBet, Lx, B)
            Er.AddMsg Msg
            GoTo Nxt
        End If
        A2.AddLBC Lx, B, C
Nxt:
    Next
'Dim O As New LABCsRslt
'Set ValidateAsBetNum = O.Init(A2, Er)
End Function

Function ValidateAsBetNumIO(ABCLines$, FnyStr$, FmNum&, ToNum&) As String()
Stop
Dim A As LABCs: Set A = LABCs.ByLines(ABCLines, IsVF:=True)
Dim R As LABCsRslt
Dim O$()
Dim Fny$(): Fny = LvsSy(FnyStr)
'Set R = A.ValidateAsBetNum(Fny, FmNum, ToNum) '<========================
   
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
Function ValidateAsFldLngVal(Fny$()) As LABCsRslt
Dim A1 As LABCsRslt: Set A1 = ValidateAsFldVal(Fny)
Stop
'Set ValidateAsFldVal = A2
End Function

Function ValidateAsFldValIO(ABCLines$, IsVF As Boolean, FnyStr$) As String()
Dim LABCAy() As LABC
Dim O$(), Fny$()
Fny = LvsSy(FnyStr)
Dim A As LABCs: Set A = LABCs.ByLines(ABCLines, IsVF)
PushAp O, "LABCs.ValidateAsFldVal '(===================="
PushAp O, "LABCs.ToStr <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", A.ToStr
PushAp O, "Oup1::ValidateAsFldVal >>>>>>>>>>>>>>>>>>>>>>", A.ValidateAsFldVal(Fny).ToStr
PushAp O, "LABCs.ValidateAsFld ')======================="
PushAp O, ""
ValidateAsFldValIO = O
End Function
Property Get ValidateAsNm() As NmRslt
Dim O As New NmRslt
Dim Er As New Er
Dim Nm$

With O
    Set .Er = Er.AddErLyAp(NmErNoLin, NmErMultiName, NmErExcessLin)
    .Nm = Nm
End With
Set ValidateAsNm = O
End Property
Function ValidateNmIO(ABCLines$)
Dim O$()
Dim A As LABCs: Set A = LABCs.ByLines(ABCLines, True)
Dim R As NmRslt: Set R = A.ValidateAsNm
PushAp O, "ValidateAsNmRslt ============================"
PushAp O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", ToStr
PushAp O, "Oup1::Nm >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.ToStr
PushAp O, "LABCAy_NmRslt ==============================="
PushAp O, ""
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

'AyDmp ValidateNmIO(InpStr)
End Sub
Friend Sub TstValidateAsFny()
Dim A As LABCs: Set A = Me.InitByT1("Lo", IsVF, "")
Debug.Print A.ToStr
Debug.Print A.ValidateAsFny.ToStr
End Sub

Private Property Get Oy() As Oy
Dim O As New Oy
Set Oy = O.Init(B_Ay)
End Property

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
Dim LABCAy() As LABC
    LABCAy = A.LABCs.Ay
Dim O As New LABCsRslt
    Dim J%, I, M As LABC
    Dim Ok As StrRslt
    If A.LABCs.IsVF Then
        Dim Lx%, B$
        For J = 0 To UB(LABCAy)
            Set Ok = VdtDupFld1(LABCAy, J)
            O.Er.Add Ok.Er
            If Ok.Str <> "" Then
                Set M = LABCAy(J)
                Lx = M.Lx
                B = M.B
                O.LABCs.AddLBC Lx, B, Ok.Str
            End If
        Next
    Else
        For J = 0 To UB(LABCAy)
            Set Ok = VdtDupFld3(LABCAy, J)
            If Ok.Str <> "" Then
                With LABCAy(J)
                    O.LABCs.AddLBC .Lx, Ok.Str, .B
                End With
            End If
            O.Er.Add Ok.Er
        Next
    End If
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
If IsVF Then
    Set VdtErFld = VdtErFld_IsVF(Fny)
Else
    Set VdtErFld = VdtErFld_NonVF(Fny)
End If
End Function

Private Function VdtErFld_NonVF(Fny$()) As LABCsRslt
Dim F%, O As New LABCs, J%, Lx%, C$, B$, FldLvs$, OEr$(), M As LABC
O.InitByT1 T1, IsVF
For J = 0 To U
    With B_Ay(J)
        Lx = .Lx
        C = .C
    End With

    VdtErFld1 Lx, C, Fny, _
        FldLvs, OEr

    If FldLvs <> "" Then
        B = B_Ay(J).B
        C = FldLvs
        O.AddLBC Lx, B, C '<============
    End If
Next
VdtErFld_NonVF = O
End Function

Private Function VdtErFld_IsVF(Fny$()) As LABCsRslt
Dim F%, O As New LABCs, J%, Lx%, C$, B$, FldLvs$, OEr$(), M As LABC
O.InitByT1 T1, IsVF
For J = 0 To U
    With B_Ay(J)
        Lx = .Lx
        C = .C
    End With

    VdtErFld1 Lx, C, Fny, _
        FldLvs, OEr

    If FldLvs <> "" Then
        B = B_Ay(J).B
        C = FldLvs
        O.AddLBC Lx, B, C '<============
    End If
Next
VdtErFld_IsVF = O
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
