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

Property Get Ay() As LABC()
Ay = B_Ay
End Property

Property Get Init(ABCAy() As ABC, Optional IsVF As Boolean) As LABCs
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
End Property

Property Get InitByLines(ABCLines$, Optional IsVF As Boolean) As LABCs
If ABCLines = "" Then PmEr
Dim Ay() As ABC, Lin
For Each Lin In SplitLines(ABCLines)
    PushObj Ay, ABC(Lin)
Next
Set InitByLines = Init(Ay, IsVF)
End Property

Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property

Property Get IsVF() As Boolean
IsVF = B_IsVF
End Property

Property Get LxFldAy() As LxFld()
Dim O() As LxFld, M As LxFld
Dim A() As LABC
    A = Me.Ay
    Dim I%, J%, FldLvs$, Lx%, Fny$()
    For J = 0 To U
        FldLvs = A(J).FldLvs
        Fny = LvsSy(FldLvs)
        Lx = A(J).Lx
        For I = 0 To UB(Fny)
            Set M = New LxFld
            M.Fld = Fny(I)
            M.Lx = Lx
            PushObj O, M
        Next
    Next
LxFldAy = O
End Property

Property Get Ly() As String()
Dim O$(), J%
For J = 0 To U
    Push O, B_Ay(J).Lin
Next
Ly = O
End Property

Property Get N&()
N = Sz(B_Ay)
End Property

Property Get T1() As String
T1 = B_T1
End Property

Property Get ToStr$()
Dim S$
    Dim O$(), J%
    For J = 0 To U
        Push O, B_Ay(J).ToStr
    Next
    S = JnCrLf(O)
ToStr = Tag("LABCs", S)
End Property

Property Get U&()
U = N - 1
End Property

Property Get UniqFny() As String()
Dim I, M As LABC, O$()
If IsEmp Then Exit Property
For Each I In B_Ay
    Set M = I
    PushNoDupAy O, M.Fny
Next
UniqFny = O
End Property

Property Get ValidateAsFny() As FnyRslt
Dim A As New LABCsRslt
Set A = VdtDupFld(A)
Dim O As New FnyRslt
Set ValidateAsFny = O.Init(UniqFny, A.Er)
End Property

Property Get ValidateAsNm() As NmRslt
Dim Nm$
    If IsEmp Then Nm = "?": Exit Function
    Dim A As LABC: Set A = B_Ay(0)
    Dim T1$: T1 = Lin(A.C).T1
    If T1 = "" Then Nm = "?": Exit Function
    Nm = T1
Dim Er As New Er
    Er.PushErLy0Ap NmErNoLin, NmErMultiName, NmErExcessLin
Dim O As New NmRslt
Set ValidateAsNm = O.Init(Nm, Er)
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

Sub AddLxFldVal(Lx%, Fld$, Val$)
If IsVF Then
    AddLBC Lx, Val, Fld
Else
    AddLBC Lx, Fld, Val
End If
End Sub

Friend Property Get CnoVals(Fny$()) As CnoVals
Dim O As New CnoVals
Dim F$, V$, Fny1$(), Cno%
Dim J%, JJ%
For J = 0 To UB(B_Ay)
    V = B_Ay(J).Val
    Fny1 = B_Ay(J).Fny
    For JJ = 0 To UB(Fny1)
        Cno = AyIx(Fny, Fny1(JJ))
        If Cno = -1 Then Stop
        Cno = Cno + 1
        O.AddCnoVal Cno, Fny1(JJ), V
    Next
Next
Set CnoVals = O
End Property

Function DupEmpLABCs() As LABCs
Dim O As New LABCs
Set DupEmpLABCs = O.InitByT1(T1, IsVF)
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

Function InitByT1(T1$, Optional IsVF As Boolean, Optional ABLy0) As LABCs
Dim ABLy$(): ABLy = DftNy(ABLy0)
If B_IsInited Then PmEr
B_IsInited = True ' Cannot init once
B_T1 = T1
B_IsVF = IsVF
Set InitByT1 = Me
End Function

Friend Property Get LxAy() As Integer()
If IsEmp Then Exit Property
Dim I, M As LABC, O%()
For Each I In B_Ay
    Set M = I
    Push O, M.Lx
Next
LxAy = O
End Property

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

Sub TstValidateAsBetNum()
Dim ABCLines$
Dim FnyStr$
Dim FmNum&
Dim ToNum&
    ABCLines = RplVBar("Wdt 10 A B C|Wdt 20 X Y Z")
AyDmp ValidateAsBetNumIO(ABCLines, FnyStr, FmNum, ToNum)
End Sub

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

Friend Sub TstValidateAsFny()
Dim A As LABCs: Set A = Me.InitByT1("Lo", IsVF, "")
Debug.Print A.ToStr
Debug.Print A.ValidateAsFny.ToStr
End Sub

Friend Sub TstValidateAsNm()
'AyDmp ValidateNmIO(InpStr)
End Sub

Function ValidateAsBetNum(Fny$(), FmNum&, ToNum&) As LABCsRslt
If Not IsVF Then PmEr
Dim A1 As LABCsRslt: Set A1 = ValidateAsFldVal(Fny)
Dim A2 As LABCsRslt: Set A2 = VdtIsNum(A1)
Set ValidateAsBetNum = VdtIsBet(A2, FmNum, ToNum)
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

Function ValidateAsFldLngVal(Fny$()) As LABCsRslt
If Not IsVF Then PmEr
Dim A1 As LABCsRslt: Set A1 = ValidateAsFldVal(Fny)
Dim A2 As LABCsRslt: Set A2 = VdtIsNum(A1)
Set ValidateAsFldLngVal = VdtIsLng(A2)
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
Dim A As LABCs: Set A = LABCs.ByLines(ABCLines, IsVF)
PushAp O, "LABCs.ValidateAsFldVal '(===================="
PushAp O, "LABCs.ToStr <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", A.ToStr
PushAp O, "Oup1::ValidateAsFldVal >>>>>>>>>>>>>>>>>>>>>>", A.ValidateAsFldVal(Fny).ToStr
PushAp O, "LABCs.ValidateAsFld ')======================="
PushAp O, ""
ValidateAsFldValIO = O
End Function

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

Private Function NmErNoLin() As String()
If IsEmp Then NmErNoLin = ApSy(FmtQQ("There is not ?-line", C2_Lo_Nam))
End Function

Private Property Get Oy() As Oy
Dim O As New Oy
Set Oy = O.Init(B_Ay)
End Property

Private Function VdtDupFld(A As LABCsRslt) As LABCsRslt
Dim Ay() As LABC
    Ay = A.LABCs.Ay

Dim LxAy%()
Dim FnyAy()
    Dim FldLvsAy$()
    LxAy = A.LABCs.LxAy
    FldLvsAy = A.LABCs.FldLvsAy
    Dim J%
    For J = 0 To UB(FldLvsAy)
        Push FnyAy, LvsSy(FldLvsAy(J))
    Next
Dim O As LABCs
Dim OEr As New Er
    OEr.Push A.Er
    Set O = A.LABCs.DupEmpLABCs
    Dim F As StrRslt ' FldLvsRslt
    For J = 0 To UB(Ay)
        Set F = VdtDupFld_1(FnyAy, LxAy, J) 'F is FnyAy(J).FldLvs after remove all duplicated fields.
                                            'If removed, F.Er will have error message
                                           '
        Set O = VdtDupFld_2(O, Ay(J), F.Str)
        OEr.Push F.Er
    Next
Set VdtDupFld = LABCsRslt(O, OEr)
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

Private Function VdtDupFld_1(FnyAy(), LxAy%(), J%) As StrRslt
'Validate Current Fny (From FnyAy(J)) has duplicate field or not
'if yes, report into Er
'return FldLvs as StrRslt after removing all the duplicated fields in Fny
Dim O As New StrRslt
    Dim OFny$()             '<-- Those Fld has no duplicated will be put in the string result
    Dim I%
    Dim F$, Fny$()
    Fny = FnyAy(J)
    For I = 0 To UB(Fny)
        F = Fny(I)
        Dim DupAtLx% ' Duplicated at which line
        With VdtDupFld_DupAtLxOpt(FnyAy, LxAy, J, I)
            If .Som Then
                Dim Lx%, Msg$
                Lx = LxAy(I)
                Msg = FmtQQ(M_Fld_IsDup, Lx, F, .Int)
                O.Er.PushMsg Msg    '<== Report duplication in OEr
            Else
                Push OFny, F    '<== Push to Fny1 for no-dup
            End If
        End With
    Next
    O.Str = JnSpc(OFny) '<== Put to {O}
Set VdtDupFld_1 = O
End Function

Private Function VdtDupFld_2(A As LABCs, M As LABC, FldLvs$) As LABCs ' (A() As LABC, FldLvs$, J%) As StrRslt
If FldLvs <> "" Then
    With M
        A.AddLxFldVal .Lx, FldLvs, .Val
    End With
End If
Set VdtDupFld_2 = A
End Function

Private Function VdtDupFld_DupAtLxOpt(FnyAy(), LxAy%(), J%, I%) As IntOpt
'Check if Fny(I)-element has duplication found in Fny(I+1..)
Dim Fny$(): Fny = FnyAy(J)
Dim F$: F = Fny(I)
Dim II%
For II = I + 1 To UB(Fny)
    If Fny(II) = F Then
        VdtDupFld_DupAtLxOpt = SomInt(II)
        Exit Function
    End If
Next
For II = J + 1 To UB(FnyAy)
    Fny = FnyAy(II)
    If AyHas(Fny, F) Then VdtDupFld_DupAtLxOpt = SomInt(II): Exit Function
Next
End Function

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
    Set O = LABCs.ByT1(T1, IsVF)
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

Private Function VdtErFld1(LvsFldAy$(), LxAy%(), Fny$()) As LyRslt
Dim OLy$(), J%, Lx%, FldLvs$, OEr As New Er, A As StrRslt
For J = 0 To UB(LvsFldAy)
    Lx = LxAy(J)
    FldLvs = LvsFldAy(J)
    Set A = VdtErFld2(FldLvs, Lx, Fny)
    Push OLy, A.Str
    OEr.Push A.Er
Next
Set VdtErFld1 = LyRslt(OLy, OEr)
End Function

Private Function VdtErFld2(LvsFld$, Lx%, Fny$()) As StrRslt
Dim F1$(): F1 = LvsSy(LvsFld)
Dim F2$(), Er As New Er
Dim J%
For J = 0 To UB(F1)
    If AyHas(Fny, F1(J)) Then
        Push F2, F1(J)
    Else
        Er.PushMsg FmtQQ(M_Fld_IsInValid, Lx, F1(J))
    End If
Next
Set VdtErFld2 = StrRslt(JnSpc(F2), Er)
End Function

Private Function VdtIsBet(A As LABCsRslt, FmNum&, ToNum&) As LABCsRslt
Dim Ay() As LABC: Ay = A.LABCs.Ay
Dim O As LABCs: Set O = A.LABCs.DupEmpLABCs
Dim OEr As New Er: OEr.Push A.Er
Dim J%, V&
For J = 0 To UB(Ay)
    With Ay(J)
        V = .B
        If FmNum <= V And V <= ToNum Then
            O.AddLBC .Lx, .B, .C
        Else
            OEr.PushMsg FmtQQ(M_Val_ShouldBet, .Lx, .B, FmNum, ToNum)
        End If
    End With
Next
Set VdtIsBet = LABCsRslt(O, OEr)
End Function

Private Function VdtIsLng(A As LABCsRslt) As LABCsRslt
Dim Ay() As LABC: Ay = A.LABCs.Ay
Dim O As LABCs: Set O = A.LABCs.DupEmpLABCs
Dim OEr As New Er: OEr.Push A.Er
Dim J%
For J = 0 To UB(Ay)
    With Ay(J)
        If IsNum(.B) Then
            O.AddLBC .Lx, .B, .C
        Else
            OEr.PushMsg FmtQQ(M_Val_IsNonLng, .Lx, .B)
        End If
    End With
Next
Set VdtIsLng = LABCsRslt(O, OEr)
End Function

Private Function VdtIsNum(A As LABCsRslt) As LABCsRslt
Dim O As LABCs: Set O = A.LABCs.DupEmpLABCs
Dim OEr As New Er: OEr.Push A.Er
    Dim Ay() As LABC: Ay = A.LABCs.Ay
    Dim J%
    For J = 0 To UB(Ay)
        With Ay(J)
            If IsNum(.B) Then
                O.AddLBC .Lx, .B, .C
            Else
                OEr.PushMsg FmtQQ(M_Val_IsNonNum, .Lx, .B)
            End If
        End With
    Next
Set VdtIsNum = LABCsRslt(O, OEr)
End Function
