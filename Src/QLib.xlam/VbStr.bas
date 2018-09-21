Attribute VB_Name = "VbStr"
Option Explicit
Type P123
    P1 As String
    P2 As String
    P3 As String
End Type
Type FmToPos
    FmPos As Long
    ToPos As Long
End Type
Type StrOpt
   Som As Boolean
   Str As String
End Type
Type LngOpt
    Som As Boolean
    Lng As Long
End Type

Function AlignL$(A, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "AlignL"
If ErIfNotEnoughWdt And DoNotCut Then
    Er CSub, "Both {ErIfNotEnoughWdt} and {DontCut} cannot be True", ErIfNotEnoughWdt, DoNotCut
End If
Dim S$: S = ValStr(A)
AlignL = StrAlignL(S, W, ErIfNotEnoughWdt, DoNotCut)
End Function

Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function

Function BktPos(A, Optional Bkt$ = "()") As FmToPos
Const CSub$ = "BktPos"
Dim Q1$, Q2$
    With BrkQuote(Bkt)
        Q1 = .S1
        Q2 = .S2
    End With
Dim FmPos&
    FmPos = InStr(A, Q1)
    If FmPos = 0 Then Exit Function
Dim ToPos&
    Dim NOpn%, J%
    For J = FmPos + 1 To Len(A)
        Select Case Mid(A, J, 1)
        Case Q2
            If NOpn = 0 Then
                ToPos = J
                Exit For
            End If
            NOpn = NOpn - 1
        Case Q1
            NOpn = NOpn + 1
        End Select
    Next
    If ToPos = 0 Then Er CSub, "{Str} has {Q1} and {Q2} is not in pair", A, Q1, Q2
'BktPos.FmPos = FmPos
'BktPos.ToPos = ToPos
End Function

Function Brk(A, Sep, Optional NoTrim As Boolean) As S1S2
Const CSub$ = "Brk"
Dim P&: P = InStr(A, Sep)
If P = 0 Then Er CSub, "{S} does not contains {Sep}", A, Sep
Brk = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Function Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Brk1 = Brk1__(A, P, Sep, NoTrim)
End Function

Function Brk1Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
Brk1Rev = Brk1__(A, P, Sep, NoTrim)
End Function

Function Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk2 = Brk2__X(A, P, Sep, NoTrim)
End Function

Sub BrkAsg(A, Sep, OS1$, OS2$, Optional NoTrim As Boolean)
With Brk(A, Sep, NoTrim)
    OS1 = .S1
    OS2 = .S2
End With
End Sub

Function BrkAt(A, P&, SepLen%, Optional NoTrim As Boolean) As S1S2
Dim O As S1S2
With O
    If NoTrim Then
        .S1 = Left(A, P - 1)
        .S2 = Mid(A, P + SepLen)
    Else
        .S1 = Trim(Left(A, P - 1))
        .S2 = Trim(Mid(A, P + SepLen))
    End If
End With
BrkAt = O
End Function

Function BrkBoth(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S1 = A
    Else
        O.S1 = Trim(A)
    End If
    O.S2 = O.S1
    BrkBoth = O
    Exit Function
End If
BrkBoth = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Function BrkByBkt(A, Optional Bkt$ = "()") As P123
Dim B As FmToPos: B = BktPos(A, Bkt)

End Function

Function BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim O As S1S2
Select Case L
Case 0:
Case 1
    O.S1 = QuoteStr
    O.S2 = O.S1
Case 2
    O.S1 = Left(QuoteStr, 1)
    O.S2 = Right(QuoteStr, 1)
Case Else
    Dim P%
    If InStr(QuoteStr, "*") > 0 Then
        O = Brk(QuoteStr, "*", NoTrim:=True)
    End If
End Select
BrkQuote = O
End Function

Function BrkRev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & A & "] does not contains Sep[" & Sep & "]"
BrkRev = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Function Dft(V, DftV)
If VarIsEmp(V) Then
   Dft = DftV
Else
   Dft = V
End If
End Function

Function DftStr(S$, DftV$)
If S = "" Then
   DftStr = DftV
Else
   DftStr = S
End If
End Function

Function Esc$(A, Fm$, ToStr$)
If InStr(A, "\n") > 0 Then
    Debug.Print ErMsgLines("Esc", "Warning: escaping a {Str} of {FmStrSub} to {ToSubStr} is found that {Str} contains some {ToSubStr}.  This will make the string chagned after UnEsc", A, Fm, ToStr)
End If
Esc = Replace(A, Fm, ToStr)
End Function

Function EscCr$(A)

End Function

Function EscCrLf$(A)
EscCrLf = EscCr(EscLf(A))
End Function

Function EscKey$(A)
EscKey = EscCrLf(EscSpc(EscTab(A)))
End Function

Function EscLf$(A)
EscLf = Esc(A, vbLf, "\n")
End Function

Function EscSpc$(A)
EscSpc = Esc(A, " ", "~")
End Function

Function EscTab$(A)
EscTab = Esc(A, vbTab, "\t")
End Function

Sub FmToPosDmp(A As FmToPos)
Debug.Print FmToPosToStr(A)
End Sub

Function FmToPosToStr$(A As FmToPos)
FmToPosToStr = FmtQQ("(FmToPos ? ?)", A.FmPos, A.ToPos)
End Function

Function FmtMacro$(MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtMacro = FmtMacroAv(MacroStr, Av)
End Function

Function FmtMacroAv$(MacroStr$, Av())
Dim Ay$(): Ay = Macro(MacroStr).Ny
Dim O$: O = MacroStr
Dim J%, I
For Each I In Ay
    O = Replace(O, I, Av(J))
    J = J + 1
Next
FmtMacroAv = O
End Function

Function FmtMacroDic$(MacroStr$, Dic As Dictionary)
Dim Ay$(): Ay = Macro(MacroStr).Ny
If Not AyIsEmp(Ay) Then
    Dim O$: O = MacroStr
    Dim I, K$
    For Each I In Ay
        K = RmvFstLasChr(CStr(I))
        If Dic.Exists(K) Then
            O = Replace(O, I, Dic(K))
        End If
    Next
End If
FmtMacroDic = O
End Function

Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function FmtQQAv$(QQVbl$, Av)
If AyIsEmp(Av) Then FmtQQAv = QQVbl: Exit Function
Dim O$
    Dim I, NeedUnEsc As Boolean
    O = RplVBar(QQVbl)
    For Each I In Av
        If InStr(I, "?") > 0 Then
            NeedUnEsc = True
            I = Replace(I, "?", Chr(255))
        End If
        O = Replace(O, "?", I, Count:=1)
    Next
    If NeedUnEsc Then O = Replace(O, Chr(255), "?")
FmtQQAv = O
End Function

Function FmtQQVBar$(QQStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQVBar = RplVBar(FmtQQAv(QQStr, Av))
End Function

Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function HasSubStr(S, SubStr) As Boolean
HasSubStr = InStr(S, SubStr) > 0
End Function

Function HasSubStrAy(S, SubStrAy) As Boolean
Dim SubStr
For Each SubStr In SubStrAy
    If HasSubStr(S, SubStr) Then HasSubStrAy = True: Exit Function
Next
End Function

Function HasVBar(S) As Boolean
HasVBar = HasSubStr(S, "|")
End Function

Function InstrN&(S, SubStr, N%)
Dim P&, J%
For J = 1 To N
    P = InStr(P + 1, S, SubStr)
    If P = 0 Then Exit Function
Next
InstrN = P
End Function

Function IsDigit(A) As Boolean
IsDigit = "0" <= A And A <= "9"
End Function

Function IsEqStrOpt(A1 As StrOpt, A2 As StrOpt) As Boolean
If A1.Som <> A1.Som Then Exit Function
If A1.Str <> A2.Str Then Exit Function
IsEqStrOpt = True
End Function

Function IsLetter(A) As Boolean
Dim C1$: C1 = UCase(A)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsNeedQuote(A) As Boolean
IsNeedQuote = True
If HasSubStr(A, " ") Then Exit Function
If HasSubStr(A, "#") Then Exit Function
If HasSubStr(A, ".") Then Exit Function
IsNeedQuote = False
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

Function IsNmChr(A) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Function
If A = "_" Then Exit Function
If IsDigit(A) Then Exit Function
IsNmChr = False
End Function

Function IsWhite(A) As Boolean
Dim B$: B = Left(A, 1)
IsWhite = True
If B = " " Then Exit Function
If B = vbCr Then Exit Function
If B = vbLf Then Exit Function
If B = vbTab Then Exit Function
IsWhite = False
End Function

Function JnComma$(Ay)
JnComma = Join(Ay, ",")
End Function

Function JnCommaSpc(Ay)
JnCommaSpc = Join(Ay, ", ")
End Function

Function JnCrLf$(Ay, Optional WithIx As Boolean)
If WithIx Then
    Dim O$(), J%
    For J = 0 To UB(Ay)
        Push O, J & ": " & Ay(J)
    Next
    JnCrLf = Join(O, vbCrLf)
Else
    JnCrLf = Join(AySy(Ay), vbCrLf)
End If
End Function

Function JnDblCrLf$(Ay)
JnDblCrLf = Join(Ay, vbCrLf & vbCrLf)
End Function

Function JnQDblComma$(Ay)
JnQDblComma = JnComma(AyQuoteDbl(Ay))
End Function

Function JnQDblSpc$(Ay)
JnQDblSpc = JnSpc(AyQuoteDbl(Ay))
End Function

Function JnQSngComma$(Ay)
JnQSngComma = JnComma(AyQuoteSng(Ay))
End Function

Function JnQSngSpc$(Ay)
JnQSngSpc = JnSpc(AyQuoteSng(Ay))
End Function

Function JnQSqBktComma$(Ay)
JnQSqBktComma = JnComma(AyQuoteSqBkt(Ay))
End Function

Function JnQSqBktSpc$(Ay)
JnQSqBktSpc = JnSpc(AyQuoteSqBkt(Ay))
End Function

Function JnSpc$(Ay)
JnSpc = Join(Ay, " ")
End Function

Function JnTab$(Ay)
JnTab = Join(Ay, vbTab)
End Function

Function JnVBar$(Ay)
JnVBar = Join(Ay, "|")
End Function

Function LTrimWhite$(A)
Dim J%
    For J = 1 To Len(A)
        If Not IsWhite(Mid(A, J, 1)) Then Exit For
    Next
LTrimWhite = Left(A, J)
End Function

Function LasChr$(A)
LasChr = Right(A, 1)
End Function

Function LvsJnComma$(Lvs$)
LvsJnComma = JnComma(LvsSy(Lvs))
End Function

Function LvsJnQuoteComma$(Lvs$)
LvsJnQuoteComma = JnComma(AyQuote(LvsSy(Lvs), "'"))
End Function

Function LvsSy(A) As String()
LvsSy = Split(RmvDblSpc(Trim(A)), " ")
End Function

Function NewFmToPos(FmPos&, ToPos&) As FmToPos
Dim O As FmToPos
With O
    .FmPos = FmPos
    .ToPos = ToPos
End With
NewFmToPos = O
End Function

Function ObjS$(A)
On Error Resume Next
ObjS = A.S
End Function

Function Prepend$(S, P)
If S <> "" Then Prepend = P & S
End Function

Function PrependDash$(S)
PrependDash = Prepend(S, "-")
End Function

Function Quote$(A, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & A & .S2
End With
End Function

Function RTrimWhite$(S)
Dim J%
    Dim A$
    For J = Len(S) To 1 Step -1
        If Not IsWhite(Mid(S, J, 1)) Then Exit For
    Next
    If J = 0 Then Exit Function
RTrimWhite = Mid(S, J)
End Function

Function Rmv2Dash$(A)
Rmv2Dash = RTrim(RmvAft(A, "--"))
End Function

Function Rmv3Dash$(A)
Rmv3Dash = RTrim(RmvAft(A, "---"))
End Function

Function RmvAft$(A, Sep$)
RmvAft = Brk1(A, Sep, NoTrim:=True).S1
End Function

Function RmvDblSpc$(A)
Dim O$: O = A
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Function

Function RmvFstLasChr$(A)
RmvFstLasChr = RmvFstChr(RmvLasChr(A))
End Function

Function RmvFstNChr$(A, Optional N% = 1)
RmvFstNChr = Mid(A, N + 1)
End Function

Function RmvLasChr$(A)
RmvLasChr = RmvLasNChr(A, 1)
End Function

Function RmvLasNChr$(A, N%)
RmvLasNChr = Left(A, Len(A) - 1)
End Function

Function RmvPfx$(S, Pfx)
Dim L%: L = Len(Pfx)
If Left(S, L) = Pfx Then
    RmvPfx = Mid(S, L + 1)
Else
    RmvPfx = S
End If
End Function

Function RmvPfxAy$(A, PfxAy)
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, CStr(Pfx)) Then RmvPfxAy = RmvPfx(A, Pfx): Exit Function
Next
RmvPfxAy = A
End Function

Function RmvSfx$(A, Sfx$)
Dim L%: L = Len(Sfx)
If Right(A, L) = Sfx Then
    RmvSfx = Left(A, Len(A) - L)
Else
    RmvSfx = A
End If
End Function

Function RplFstChr$(A, By$)
RplFstChr = By & RmvFstChr(A)
End Function

Function RplPfx(A, FmPfx, ToPfx)
RplPfx = ToPfx & RmvPfx(A, FmPfx)
End Function

Function RplQ$(A, By$)
RplQ = Replace(A, "?", By)
End Function

Function RplVBar$(A)
RplVBar = Replace(A, "|", vbCrLf)
End Function

Function S1S2_IsEmp(A() As S1S2) As Boolean
S1S2_IsEmp = S1S2_Sz(A) = 0
End Function

Sub S1S2_Push(O() As S1S2, M As S1S2)
Dim N&
    N = S1S2_Sz(O)
ReDim Preserve O(N)
    O(N) = M
End Sub

Function SomLng(A&) As LngOpt
With SomLng
    .Som = True
    .Lng = A
End With
End Function

Function SplitComma(A, Optional NoTrim As Boolean) As String()
If NoTrim Then
    SplitComma = Split(A, ",")
Else
    SplitComma = AyTrim(Split(A, ","))
End If
End Function

Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function

Function SplitLf(A) As String()
SplitLf = Split(A, vbLf)
End Function

Function SplitLines(A) As String()
Dim B$: B = Replace(A, vbCrLf, vbLf)
SplitLines = SplitLf(B)
End Function

Function SplitSpc(A) As String()
SplitSpc = Split(A, " ")
End Function

Function SplitVBar(A, Optional Trim As Boolean) As String()
If Trim Then
    SplitVBar = AyTrim(Split(A, "|"))
Else
    SplitVBar = Split(A, "|")
End If
End Function

Function StrAlignL$(S$, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "StrAlignL"
Dim L%: L = Len(S)
If L > W Then
    If ErIfNotEnoughWdt Then
        Er CSub, "Len({S)) > {W}", S, W
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

Sub StrBrw(A, Optional Fnn$)
Dim T$: T = TmpFt("StrBrw", Fnn$)
StrWrt A, T
FtBrw T
End Sub

Function StrDup$(N%, S)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function StrOptAy_HasNone(A() As StrOpt) As Boolean
Dim J%
For J = 0 To StrOpt_UB(A)
    If Not A(J).Som Then StrOptAy_HasNone = True: Exit Function
Next
End Function

Sub StrOpt_Dmp(A As StrOpt)
Debug.Print StrOpt_Str(A)
End Sub

Sub StrOpt_Push(O() As StrOpt, A As StrOpt)
Dim N&: N = StrOpt_Sz(O)
End Sub

Function StrOpt_Str$(A As StrOpt, Optional W% = 50)
With A
    If .Som Then
        If Len(A.Str) < W Then
            StrOpt_Str = "*SomStr " & A.Str
        Else
            StrOpt_Str = "*SomStr " & AlignL(A.Str, 50)
        End If
    Else
        StrOpt_Str = "*NoStr"
    End If
End With
End Function

Function StrOpt_Sz&(A() As StrOpt)
On Error Resume Next
StrOpt_Sz = UBound(A) - 1
End Function

Function StrOpt_UB&(A() As StrOpt)
StrOpt_UB = StrOpt_Sz(A) - 1
End Function

Function StrPfx$(A, PfxAy$())
If AyIsEmp(PfxAy) Then Exit Function
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, CStr(Pfx)) Then StrPfx = Pfx: Exit Function
Next
End Function

Sub StrWrt(A, Ft)
Fso.CreateTextFile(Ft, True).Write A
End Sub

Function SubStrCnt&(A, SubStr)
Dim P&: P = 1
Dim L%: L = Len(SubStr)
Dim O%
While P > 0
    P = InStr(P, A, SubStr)
    If P = 0 Then SubStrCnt = O: Exit Function
    O = O + 1
    P = P + L
Wend
SubStrCnt = O
End Function

Function SubStrPos(A, SubStr$) As FmToPos
Dim FmPos&: FmPos = InStr(A, SubStr)
Dim ToPos&
If FmPos > 0 Then ToPos = FmPos + Len(SubStr)
SubStrPos = NewFmToPos(FmPos, ToPos)
End Function

Function TrimWhite$(A)
TrimWhite = RTrimWhite(LTrimWhite(A))
End Function

Function UnEscSpc$(A)
UnEscSpc = Replace(A, "~", " ")
End Function

Function UnEscTab(A)
UnEscTab = Replace(A, "\t", "~")
End Function

Function ValStr$(A)
If VarIsPrim(A) Then ValStr = A: Exit Function
If IsNothing(A) Then ValStr = "#Nothing": Exit Function
If IsEmpty(A) Then ValStr = "#Empty": Exit Function
Dim T$
If IsObject(A) Then
    T = TypeName(A)
    Select Case T
    Case "CodeModule"
        ValStr = FmtQQ("*Md{?}", CvMdx(A).Nm)
        Exit Function
    End Select
    ValStr = FmtQQ("*?{?}", T, ObjS(A))
    Exit Function
End If
If IsArray(A) Then
    ValStr = "*Array"
    Exit Function
End If
Stop
End Function

Function VarLngOpt(V) As LngOpt
Dim O&
On Error GoTo X
O = V
VarLngOpt = SomLng(O)
Exit Function
X:
End Function

Private Sub Brk1Rev__Tst()
Dim S1$, S2$, ExpS1$, ExpS2$, A$
A = "aa --- bb --- cc"
ExpS1 = "aa --- bb"
ExpS2 = "cc"
With Brk1Rev(A, "---")
    S1 = .S1
    S2 = .S2
End With
Ass S1 = ExpS1
Ass S2 = ExpS2
End Sub

Private Function Brk1__(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S1 = A
    Else
        O.S1 = Trim(A)
    End If
    Brk1__ = O
    Exit Function
End If
Brk1__ = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Private Sub InstrN__Tst()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InstrN(S, SubStr, N)
Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InstrN(S, SubStr, N)
Ass Exp = Act
End Sub

Private Sub RmvPfx__Tst()
Ass RmvPfx("aaBB", "aa") = "BB"
End Sub

Sub RplPfx__Tst()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub

Function SubStrCnt__Tst()
Ass SubStrCnt("aaaa", "aa") = 2
Ass SubStrCnt("aaaa", "a") = 4
End Function
