Attribute VB_Name = "F_LABC"
Option Explicit
'Enum eTstOpt
'    eAllValidate = 0
'    eValidateAsFldVal = 1
'    eValidateAsBetNum = 2
'    eValidateAsNm = 3
'    eValidateAsFny = 4
'End Enum
'
'Property Get LABCAy_CnoValAy(A() As LABC, Fny$()) As CnoVal()
'Dim O() As CnoVal
'Dim F$, V$, Fny1$(), Cno%
'Dim J%, JJ%
'For J = 0 To UB(A)
'Stop '
''    V = A(J).Val
''    Fny1 = A(J).Fny
'    For JJ = 0 To UB(Fny1)
'        Cno = AyIx(Fny, Fny1(JJ))
'        If Cno = -1 Then Stop
'        Cno = Cno + 1
''        O.AddCnoVal Cno, Fny1(JJ), V
'    Next
'Next
'LABCAy_CnoValAy = O
'End Property
'
'
'Sub AddLBC(Lx%, B$, C$)
'Dim O As New LABC
'With O
'    .Lx = Lx
'    .B = B
'    .C = C
'End With
'Push A, O
'End Sub
'
'Property Get LFVAy_FnyRslt(A As LABCGp) As FnyRslt
'Dim O As New LABCAyRslt
'Set A = VdtDupFld(A)
'Dim O As New FnyRslt
'Set LABCAy_ValidateAsFny = O.Init(UniqFny, A.Er, C1, C2)
'End Property
'
'Property Get Init(ABCAy() As ABC, Optional IsVF As Boolean) As LABC()
'If B_IsInited Then PmEr
'B_IsInited = True ' Cannot init once
'If AyIsEmp(ABCAy) Then PmEr
'If Not AyIsAllEq(Oy.PrpSy("A")) Then PmEr
'B_T1 = ABCAy(0).A
'Dim ABC As ABC, I, Lx%
'For Each I In ABCAy
'    Set ABC = I
'    With ABC
'        AddLBC Lx, .B, .C
'    End With
'    Lx = Lx + 1
'Next
'Set Init = Me
'End Property
'
'Property Get InitByLines(ABCLines$, Optional IsVF As Boolean) As LABC()
'If ABCLines = "" Then PmEr
'Dim Ay() As ABC, Lin
'For Each Lin In SplitLines(ABCLines)
'    PushObj Ay, ABC(Lin)
'Next
'Set InitByLines = Init(Ay, IsVF)
'End Property
'
'Property Get InitByT1(T1$, Optional IsVF As Boolean, Optional ABLy0) As LABC()
'Dim ABLy$(): ABLy = DftNy(ABLy0)
'If B_IsInited Then PmEr
'B_IsInited = True ' Cannot init once
'B_T1 = T1
'B_IsVF = IsVF
'Set InitByT1 = Me
'End Property
'
'
'Private Property Get NmErNoLin() As String()
'Stop
''If IsEmp Then NmErNoLin = ApSy(FmtQQ("There is not ?-line", C2_Lo_Nam))
'End Property
'
'Private Property Get NmErExcessLin() As String()
'Dim J%, O$()
'If N <= 1 Then Exit Property
'For J = 1 To U
'    Push O, FmtQQ(M_Nm_ExcessLin, A(J).Lx)
'Next
'NmErExcessLin = O
'End Property
'
'Private Property Get NmErMultiName() As String()
'If IsEmp Then Exit Property
'Dim A As LABC: Set A = A(0)
'End Property
'
'Property Get Ly() As String()
'Dim O$(), J%
'For J = 0 To U
'    Push O, A(J).Lin
'Next
'Ly = O
'End Property
'
'Property Get ToStr$()
'Dim S$
'    Dim O$(), J%
'    For J = 0 To U
'        Push O, A(J).ToStr
'    Next
'    S = JnCrLf(O)
'ToStr = Tag("LABC()", S)
'End Property
'
'Sub TstValidateAsBetNum()
'Dim ABCLines$
'Dim FnyStr$
'Dim FmNum&
'Dim ToNum&
'    ABCLines = RplVBar("Wdt 10 A B C|Wdt 20 X Y Z")
'AyDmp ValidateAsBetNumIO(ABCLines, FnyStr, FmNum, ToNum)
'End Sub
'
'Property Get LxFldAy() As LxFld()
'Dim O() As LxFld, M As LxFld
'Dim A() As LABC
'    A = Me.Ay
'    Dim I%, J%, FldSsl$, Lx%, Fny$()
'    For J = 0 To U
'        FldSsl = A(J).FldSsl
'        Fny = SslSy(FldSsl)
'        Lx = A(J).Lx
'        For I = 0 To UB(Fny)
'            Set M = New LxFld
'            M.Fld = Fny(I)
'            M.Lx = Lx
'            PushObj O, M
'        Next
'    Next
'LxFldAy = O
'End Property
'
'Property Get UniqFny() As String()
'Dim I, M As LABC, O$()
'If IsEmp Then Exit Property
'For Each I In A
'    Set M = I
'    PushNoDupAy O, M.Fny
'Next
'UniqFny = O
'End Property
'
'Property Get ValidateAsBetNum(Fny$(), FmNum&, ToNum&) As LABCAyRslt
'If Not IsVF Then PmEr
'Dim A1 As LABCAyRslt: Set A1 = ValidateAsFldVal(Fny)
'Dim A2 As LABCAyRslt: Set A2 = VdtIsNum(A1)
'Set ValidateAsBetNum = VdtIsBet(A2, FmNum, ToNum)
'End Property
'
'Private Property Get VdtIsBet(A As LABCAyRslt, FmNum&, ToNum&) As LABCAyRslt
'Dim Ay() As LABC: Ay = A.LABC().Ay
'Dim O As LABC(): Set O = A.LABC().DupEmpLABC()
'Dim OEr As New Er: OEr.Push A.Er
'Dim J%, V&
'For J = 0 To UB(Ay)
'    With Ay(J)
'        V = .B
'        If FmNum <= V And V <= ToNum Then
'            O.AddLBC .Lx, .B, .C
'        Else
'            OEr.PushMsg FmtQQ(M_Val_ShouldBet, .Lx, .B, FmNum, ToNum)
'        End If
'    End With
'Next
'Set VdtIsBet = LABCAyRslt(O, OEr)
'End Property
'
'Property Get ValidateAsBetNumIO(ABCLines$, FnyStr$, FmNum&, ToNum&) As String()
'Stop
'Dim A As LABC(): Set A = LABC().ByLines(ABCLines, IsVF:=True)
'Dim R As LABCAyRslt
'Dim O$()
'Dim Fny$(): Fny = SslSy(FnyStr)
''Set R = A.ValidateAsBetNum(Fny, FmNum, ToNum) '<========================
'
'PushAp O, "LABCAy_LCFVRsltOfBetNum======================"
'PushAp O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", ToStr
'PushAp O, "Inp2::Fny <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", FnyStr
'PushAp O, "Inp3::FmNum ToNum <<<<<<<<<<<<<<<<<<<<<<<<<<<", FmtQQ("FmToNum(? ?)", FmNum, ToNum)
'PushAp O, "Oup1::Ok >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.ToStr
'PushAp O, "LABCAy_LCFVRsltOfBetNum======================"
'PushAp O, ""
'ValidateAsBetNumIO = O
'End Property
'
'Property Get ValidateAsFldVal(Fny$()) As LABCAyRslt
'Dim A1 As LABCAyRslt: Set A1 = VdtErFld(Fny)
'Dim A2 As LABCAyRslt: Set A2 = VdtDupFld(A1)
'Set ValidateAsFldVal = A2
'End Property
'
'Property Get ValidateAsFldLngVal(Fny$()) As LABCAyRslt
'If Not IsVF Then PmEr
'Dim A1 As LABCAyRslt: Set A1 = ValidateAsFldVal(Fny)
'Dim A2 As LABCAyRslt: Set A2 = VdtIsNum(A1)
'Set ValidateAsFldLngVal = VdtIsLng(A2)
'End Property
'
'Private Property Get VdtIsNum(A As LABCAyRslt) As LABCAyRslt
'Dim O As LABC(): Set O = A.LABC().DupEmpLABC()
'Dim OEr As New Er: OEr.Push A.Er
'    Dim Ay() As LABC: Ay = A.LABC().Ay
'    Dim J%
'    For J = 0 To UB(Ay)
'        With Ay(J)
'            If IsNumeric(.B) Then
'                O.AddLBC .Lx, .B, .C
'            Else
'                OEr.PushMsg FmtQQ(M_Val_IsNonNum, .Lx, .B)
'            End If
'        End With
'    Next
'Set VdtIsNum = LABCAyRslt(O, OEr)
'End Property
'
'Private Property Get VdtIsLng(A As LABCAyRslt) As LABCAyRslt
'Dim Ay() As LABC: Ay = A.LABC().Ay
'Dim O As LABC(): Set O = A.LABC().DupEmpLABC()
'Dim OEr As New Er: OEr.Push A.Er
'Dim J%
'For J = 0 To UB(Ay)
'    With Ay(J)
'        If IsLng(.B) Then
'            O.AddLBC .Lx, .B, .C
'        Else
'            OEr.PushMsg FmtQQ(M_Val_IsNonLng, .Lx, .B)
'        End If
'    End With
'Next
'Set VdtIsLng = LABCAyRslt(O, OEr)
'End Property
'
'Property Get ValidateAsFldValIO(ABCLines$, IsVF As Boolean, FnyStr$) As String()
'Dim LABCAy() As LABC
'Dim O$(), Fny$()
'Fny = SslSy(FnyStr)
'Dim A As LABC(): Set A = LABC().ByLines(ABCLines, IsVF)
'PushAp O, "LABC().ValidateAsFldVal '(===================="
'PushAp O, "LABC().ToStr <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", A.ToStr
'PushAp O, "Oup1::ValidateAsFldVal >>>>>>>>>>>>>>>>>>>>>>", A.ValidateAsFldVal(Fny).ToStr
'PushAp O, "LABC().ValidateAsFld ')======================="
'PushAp O, ""
'ValidateAsFldValIO = O
'End Property
'
'Property Get ValidateAsNm() As NmRslt
'Dim Nm$
'    If IsEmp Then Nm = "?": Exit Function
'    Dim A As LABC: Set A = A(0)
'    Dim T1$: T1 = LinT1(A.C)
'    If T1 = "" Then Nm = "?": Exit Function
'    Nm = T1
'Dim Er As New Er
'    Er.PushErLy0Ap NmErNoLin, NmErMultiName, NmErExcessLin
'Dim O As New NmRslt
'Set ValidateAsNm = O.Init(Nm, Er)
'End Property
'
'Property Get ValidateNmIO(ABCLines$)
'Dim O$()
'Dim A As LABC(): Set A = LABC().ByLines(ABCLines, True)
'Dim R As NmRslt: Set R = A.ValidateAsNm
'PushAp O, "ValidateAsNmRslt ============================"
'PushAp O, "Inp1::LABCAy <<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<<", ToStr
'PushAp O, "Oup1::Nm >>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>>", R.ToStr
'PushAp O, "LABCAy_NmRslt ==============================="
'PushAp O, ""
'ValidateNmIO = O
'End Property
'
'Friend Sub TstValidateAsFldVal()
'Dim IsVF As Boolean, ABCLines$, FnyStr$
'IsVF = True
'ABCLines = _
'    "Wdt 10 A B C X" & vbCrLf & _
'    "Wdt 20 A B D Y A"
'FnyStr = "A B C D E X"
'Debug.Print "LABC().TstValidateAsFldVal"
'AyDmp ValidateAsFldValIO(ABCLines, IsVF, FnyStr)
'End Property
'
'Sub TstValidateAsNm()
''AyDmp ValidateNmIO(InpStr)
'End Sub
'
'Sub TstValidateAsFny()
'Dim A As LABC(): Set A = Me.InitByT1("Lo", IsVF, "")
'Debug.Print A.ToStr
'Debug.Print A.ValidateAsFny("Lo", "Fny").ToStr
'End Sub
'
'Private Property Get Oy() As Oy
'Dim O As New Oy
'Set Oy = O.Init(A)
'End Property
'
'Private Property Get VdtDupFld_1(FnyAy(), LxAy%(), J%) As StrRslt
''Validate Current Fny (From FnyAy(J)) has duplicate field or not
''if yes, report into Er
''return FldSsl as StrRslt after removing all the duplicated fields in Fny
'Dim O As New StrRslt
'    Dim OFny$()             '<-- Those Fld has no duplicated will be put in the string result
'    Dim I%
'    Dim F$, Fny$()
'    Fny = FnyAy(J)
'    For I = 0 To UB(Fny)
'        F = Fny(I)
'        Dim DupAtLx% ' Duplicated at which line
'        With VdtDupFld_DupAtLxOpt(FnyAy, LxAy, J, I)
'            If .Som Then
'                Dim Lx%, Msg$
'                Lx = LxAy(I)
'                Msg = FmtQQ(M_Fld_IsDup, Lx, F, .I)
'                O.Er.PushMsg Msg    '<== Report duplication in OEr
'            Else
'                Push OFny, F    '<== Push to Fny1 for no-dup
'            End If
'        End With
'    Next
'    O.Str = JnSpc(OFny) '<== Put to {O}
'Set VdtDupFld_1 = O
'End Property
'
'Private Property Get VdtDupFld_DupAtLxOpt(FnyAy(), LxAy%(), J%, I%) As SomInt
''Check if Fny(I)-element has duplication found in Fny(I+1..)
'Dim Fny$(): Fny = FnyAy(J)
'Dim F$: F = Fny(I)
'Dim II%
'For II = I + 1 To UB(Fny)
'    If Fny(II) = F Then
'        VdtDupFld_DupAtLxOpt = SomInt(II)
'        Exit Property
'    End If
'Next
'For II = J + 1 To UB(FnyAy)
'    Fny = FnyAy(II)
'    If AyHas(Fny, F) Then VdtDupFld_DupAtLxOpt = SomInt(II): Exit Property
'Next
'End Property
'
'Private Property Get VdtDupFld2(A() As LABC, J%, F$, _
'    ODupAtLx%) As Boolean
''Check if F has duplicated-element found in A(J+1...)
'Dim JJ%, Fny$(), FldSsl$
'ODupAtLx = -1
''For JJ = J + 1 To LABC_UB(A)
''    FldSsl = A(JJ).C
''    Fny = SslSy(FldSsl)
''    If AyHas(Fny, F) Then
''        ODupAtLx = JJ
''        VdtDupFld2 = True
''        Exit Function
''    End If
''Next
'End Property
'
'Private Property Get VdtDupFld(A As LABCAyRslt) As LABCAyRslt
'Dim Ay() As LABC
'    Ay = A.LABC().Ay
'
'Dim LxAy%()
'Dim FnyAy()
'    Dim FldSslAy$()
'    LxAy = A.LABC().LxAy
'    FldSslAy = A.LABC().FldSslAy
'    Dim J%
'    For J = 0 To UB(FldSslAy)
'        Push FnyAy, SslSy(FldSslAy(J))
'    Next
'Dim O As LABC()
'Dim OEr As New Er
'    OEr.Push A.Er
'    Set O = A.LABC().DupEmpLABC()
'    Dim F As StrRslt ' FldSslRslt
'    For J = 0 To UB(Ay)
'        Set F = VdtDupFld_1(FnyAy, LxAy, J) 'F is FnyAy(J).FldSsl after remove all duplicated fields.
'                                            'If removed, F.Er will have error message
'                                           '
'        Set O = VdtDupFld_2(O, Ay(J), F.Str)
'        OEr.Push F.Er
'    Next
'Set VdtDupFld = LABCAyRslt(O, OEr)
'End Property
'
'Private Property Get VdtDupFld_2(A As LABC(), M As LABC, FldSsl$) As LABC() ' (A() As LABC, FldSsl$, J%) As StrRslt
'If FldSsl <> "" Then
'    With M
'        A.AddLxFldVal .Lx, FldSsl, .Val
'    End With
'End If
'Set VdtDupFld_2 = A
'End Property
'
'Friend Property Get FldSslAy() As String()
'If IsEmp Then Exit Property
'Dim I, M As LABC, O$()
'For Each I In A
'    Set M = I
'    Push O, M.FldSsl
'Next
'FldSslAy = O
'End Property
'
'Friend Property Get LxAy() As Integer()
'If IsEmp Then Exit Property
'Dim I, M As LABC, O%()
'For Each I In A
'    Set M = I
'    Push O, M.Lx
'Next
'LxAy = O
'End Property
'
'Private Property Get VdtErFld(Fny$()) As LABCAyRslt
'Dim Ly$(), LxAy%()
'LxAy = Me.LxAy
'Dim Er As Er
'    With VdtErFld1(Me.FldSslAy, LxAy, Fny)
'        Set Er = .Er
'        Ly = .Ly
'    End With
'Dim O As New LABC()
'    Dim J%, B$, C$, Lx%
'    Set O = LABC().ByT1(T1, IsVF)
'    If IsVF Then
'        For J = 0 To UB(Ly)
'            Lx = LxAy(J)
'            B = A(J).B
'            C = Ly(J)
'            O.AddLBC Lx, B, C
'        Next
'    Else
'        For J = 0 To UB(Ly)
'            Lx = LxAy(J)
'            C = A(J).C
'            B = Ly(J)
'            O.AddLBC Lx, B, C
'        Next
'    End If
'Set VdtErFld = LABCAyRslt(O, Er)
'End Property
'
'Private Property Get VdtErFld2(SslFld$, Lx%, Fny$()) As StrRslt
'Dim F1$(): F1 = SslSy(SslFld)
'Dim F2$(), Er As New Er
'Dim J%
'For J = 0 To UB(F1)
'    If AyHas(Fny, F1(J)) Then
'        Push F2, F1(J)
'    Else
'        Er.PushMsg FmtQQ(M_Fld_IsInValid, Lx, F1(J))
'    End If
'Next
'Set VdtErFld2 = StrRslt(JnSpc(F2), Er)
'End Property
'
'Private Property Get VdtErFld1(SslFldAy$(), LxAy%(), Fny$()) As LyRslt
'Dim OLy$(), J%, Lx%, FldSsl$, OEr As New Er, A As StrRslt
'For J = 0 To UB(SslFldAy)
'    Lx = LxAy(J)
'    FldSsl = SslFldAy(J)
'    Set A = VdtErFld2(FldSsl, Lx, Fny)
'    Push OLy, A.Str
'    OEr.Push A.Er
'Next
'Set VdtErFld1 = LyRslt(OLy, OEr)
'End Property
'
'Sub Tst(Optional Opt As eTstOpt = eValidateAsFldVal)
'Select Case Opt
'Case eValidateAsFldVal: TstValidateAsFldVal
'Case eValidateAsBetNum: TstValidateAsBetNum
'Case eValidateAsNm:     TstValidateAsNm
'Case eValidateAsFny:    TstValidateAsFny
'Case eAllValidate:
'    TstValidateAsFldVal
'    TstValidateAsBetNum
'    TstValidateAsNm
'    TstValidateAsFny
'Case Else
'    PmEr
'End Select
'End Sub
'
'Property Get Ly_LABCAy(ABCAy() As ABC, IsVF As Boolean) As LABC()
'Dim O As New LABC()
'Set Init = O.Init(ABCAy, IsVF)
'End Property
'
'Property Get ABCLines_LABCAy(ABCLines$, Optional IsVF As Boolean) As LABC()
'Dim O As New LABC()
'Set ByLines = O.InitByLines(ABCLines, IsVF)
'End Property
'
'Property Get ABCVbl_LABCAy(ABCVBarLines$, Optional IsVF As Boolean) As LABC()
'Set ByVBarLines = ByLines(RplVBar(ABCVBarLines), IsVF)
'End Property
'
'
