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
Enum eTstOpt
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
'O.SetCnoAyWdtAy .CnoAy, AyIntAy(.ValAy)
Stop
End With
Set CnoAyWdtAy = O
End Property
Private Property Get CnoAyValAy(Fny$()) As CnoAyValAy
Dim OCno%(), OVal()
Dim F$, V$, Fny1$(), Cno%
Dim J%, JJ%
For J = 0 To UB(B_Ay)
    V = B_Ay(J).Val
    Fny1 = B_Ay(J).Fny
    For JJ = 0 To UB(Fny1)
        Cno = AyIx(Fny, Fny1(JJ))
        Push OCno, Cno
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
If B_IsInited Then PmEr
B_IsInited = True ' Cannot init once
If AyIsEmp(ABCAy) Then PmEr
If Not AyIsAllEq(Oy.PrpSy("A")) Then PmEr
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
If ABCLines = "" Then PmEr
Dim Ay() As ABC, Lin
For Each Lin In SplitLines(ABCLines)
    PushObj Ay, ABC(Lin)
Next
Set InitByLines = Init(Ay, IsVF)
End Function

Function InitByT1(T1$, Optional IsVF As Boolean, Optional ABLy0) As LABCs
Dim ABLy$(): ABLy = DftNy(ABLy0)
If B_IsInited Then PmEr
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
Set ValidateAsBetNum = LABCsRslt(A2, Er)
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
Set ValidateAsFldLngVal = A1
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

Private Function VdtDupFld_FldLvsRslt(FldLvsAy$(), LxAy%(), J%) As StrRslt
'Validate Current FldLvsAy(J) has duplicate field or not
'if yes, report into Er
'if no,  add to Fny1 so that Fny1 will be JnSpc as str rslt
Dim Fny$()
    Dim FldLvs$
    FldLvs = FldLvsAy(J)
    Fny = LvsSy(FldLvs) '<--- All Fld to be checked

Dim O As New StrRslt
    Dim OFny$()             '<-- Those Fld has no duplicated will be put in the string result
    Dim I%
    Dim F$
    For I = 0 To UB(Fny)
        F = Fny(I)
        Dim DupAtLx% ' Duplicated at which line
        With VdtDupFld_Fld_I_IsDupAtLxOpt(Fny, I)
            If .Som Then
                Dim Lx%, Msg$
                Lx = LxAy(I)
                Msg = FmtQQ(M_Fld_IsDup, Lx, F, .Int)
                O.Er.AddMsg Msg    '<== Report duplication in OEr
            Else
                Push OFny, F    '<== Push to Fny1 for no-dup
            End If
        End With
    Next
    O.Str = JnSpc(OFny) '<== Put to {O}
Set VdtDupFld_FldLvsRslt = O
End Function

Private Function VdtDupFld_Fld_I_IsDupAtLxOpt(Fny$(), I%) As IntOpt
'Check if Fny(I)-element has duplication found in Fny(I+1..)
Dim F$: F = Fny(I)
Dim II%
For II = I + 1 To UB(Fny)
    If Fny(II) = F Then VdtDupFld_Fld_I_IsDupAtLxOpt = SomInt(II): Exit Function
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
Dim IsVF As Boolean
Dim LxAy%()
Dim FldLvsAy$()
    LxAy = A.LABCs.LxAy
    IsVF = A.LABCs.IsVF
    FldLvsAy = A.LABCs.FldLvsAy
Dim O As New LABCsRslt
Dim J%, I, M As LABC, FldLvs$
Dim F As StrRslt ' FldLvsRslt
Dim Lx%, B$, Ay() As LABC
Ay = A.LABCs.Ay
For J = 0 To UB(Ay)
    Set F = VdtDupFld_FldLvsRslt(FldLvsAy, LxAy, J)
    O.Er.Add F.Er
    If F.Str <> "" Then
        Set M = Ay(J)
        Lx = M.Lx
        B = M.B
        O.LABCs.AddLBC Lx, B, F.Str
    End If
Next
Set VdtDupFld = O
End Function

Private Function VdtDupFld_FldLvsRslt1() ' (A() As LABC, FldLvs$, J%) As StrRslt
Dim Fny1$(), Fny2$(), I%, F$, Msg$, O As New StrRslt
'Fny1 = LvsSy(FldLvs)
For I = 0 To UB(Fny1)
'    F = Fny1(J)
'    If Not AyHas(Fny, F) Then
'        Msg = FmtQQ(M_Fld_IsInValid, Lx, F)
'        OEr.AddMsg Msg   '<=================
'    Else
        Push Fny2, F
'    End If
Next
O.Str = JnSpc(Fny2)

'Set VdtDupFld_FldLvsRslt = O
End Function

Friend Property Get FldLvsAy() As String()
If IsEmp Then Exit Property
Dim I, M As LABC, O$()
For Each I In B_Ay
    Set M = I
    Push O, M.FldLvs
Next
FldLvsAy = O
End Property
Friend Property Get LxAy() As Integer()
If IsEmp Then Exit Property
Dim I, M As LABC, O%()
For Each I In B_Ay
    Set M = I
    Push O, M.Lx
Next
LxAy = O
End Property

Private Function VdtErFld(Fny$()) As LABCsRslt
Dim Ly$(), LxAy%()
LxAy = Me.LxAy
Dim Er As Er
    With VdtErFld1(Me.FldLvsAy, LxAy, Fny)
        Set Er = .Er
        Ly = .Ly
    End With
Dim O As New LABCs
    Dim J%, B$, C$, Lx%
    O.InitByT1 T1, IsVF
    If IsVF Then
        For J = 0 To UB(Ly)
            Lx = LxAy(J)
            B = B_Ay(J).B
            C = Ly(J)
            O.AddLBC Lx, B, C
        Next
    Else
        For J = 0 To UB(Ly)
            Lx = LxAy(J)
            C = B_Ay(J).C
            B = Ly(J)
            O.AddLBC Lx, B, C
        Next
    End If
Set VdtErFld = LABCsRslt(O, Er)
End Function

Private Function VdtErFld2(LvsFld$, Lx%, Fny$()) As StrRslt
Dim F1$(): F1 = LvsSy(LvsFld)
Dim F2$(), Er As New Er
Dim J%
For J = 0 To UB(F1)
    If AyHas(Fny, F1(J)) Then
        Push F2, F1(J)
    Else
        Er.AddMsg FmtQQ(M_Fld_IsInValid, Lx, F1(J))
    End If
Next
Set VdtErFld2 = StrRslt(JnSpc(F2), Er)
End Function

Private Function VdtErFld1(LvsFldAy$(), LxAy%(), Fny$()) As LyRslt
Dim OLy$(), J%, Lx%, FldLvs$, OEr As New Er, A As StrRslt
For J = 0 To UB(LvsFldAy)
    Lx = LxAy(J)
    FldLvs = LvsFldAy(J)
    Set A = VdtErFld2(FldLvs, Lx, Fny)
    Push OLy, A.Str
    OEr.Add A.Er
Next
Set VdtErFld1 = LyRslt(OLy, OEr)
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
    PmEr
End Select
End Sub
