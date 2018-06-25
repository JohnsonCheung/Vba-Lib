Attribute VB_Name = "M_Str"
Option Explicit

Property Get AlignL$(A, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "AlignL"
If W <= 0 Then AlignL = A: Exit Property
If ErIfNotEnoughWdt And DoNotCut Then
    Stop
    'Er CSub, "Both {ErIfNotEnoughWdt} and {DontCut} cannot be True", ErIfNotEnoughWdt, DoNotCut
End If
Dim S$: S = ToStr(A)
AlignL = StrAlignL(S, W, ErIfNotEnoughWdt, DoNotCut)
End Property

Property Get AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Property

Property Get AyTrim(A) As String()
If AyIsEmp(A) Then Exit Property
Dim U&
    U = UB(A)
Dim O$()
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = Trim(A(J))
    Next
AyTrim = O
End Property

Property Get FstChr$(A)
FstChr = Left(A, 1)
End Property

Property Get HasSubStr(S, SubStr) As Boolean
HasSubStr = InStr(S, SubStr) > 0
End Property

Property Get HasSubStrAy(S, SubStrAy) As Boolean
Dim SubStr
For Each SubStr In SubStrAy
    If HasSubStr(S, SubStr) Then HasSubStrAy = True: Exit Property
Next
End Property

Property Get HasVBar(S) As Boolean
HasVBar = HasSubStr(S, "|")
End Property

Property Get InstrN&(S, SubStr, N%)
Dim P&, J%
For J = 1 To N
    P = InStr(P + 1, S, SubStr)
    If P = 0 Then Exit Property
Next
InstrN = P
End Property

Property Get LTrimWhite$(A)
Dim J%
    For J = 1 To Len(A)
        If Not IsWhite(Mid(A, J, 1)) Then Exit For
    Next
LTrimWhite = Left(A, J)
End Property

Property Get LasChr$(A)
LasChr = Right(A, 1)
End Property

Property Get Prepend$(S, P)
If S <> "" Then Prepend = P & S
End Property

Property Get PrependDash$(S)
PrependDash = Prepend(S, "-")
End Property

Property Get RTrimWhite$(S)
Dim J%
    Dim A$
    For J = Len(S) To 1 Step -1
        If Not IsWhite(Mid(S, J, 1)) Then Exit For
    Next
    If J = 0 Then Exit Property
RTrimWhite = Mid(S, J)
End Property

Property Get SplitSpc(A) As String()
SplitSpc = Split(A, " ")
End Property

Property Get SplitVBar(A, Optional Trim As Boolean) As String()
If Trim Then
    SplitVBar = AyTrim(Split(A, "|"))
Else
    SplitVBar = Split(A, "|")
End If
End Property

Property Get StrAlignL$(S$, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "StrAlignL"
Dim L%: L = Len(S)
If L > W Then
    If ErIfNotEnoughWdt Then
        Stop
        'Er CSub, "Len({S)) > {W}", S, W
    End If
    If DoNotCut Then
        StrAlignL = S
        Exit Property
    End If
End If

If W >= L Then
    StrAlignL = S & Space(W - L)
    Exit Property
End If
If W > 2 Then
    StrAlignL = Left(S, W - 2) + ".."
    Exit Property
End If
StrAlignL = Left(S, W)
End Property

Property Get StrDup$(N%, S)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Property

Property Get StrPfx$(A, PfxAy$())
If AyIsEmp(PfxAy) Then Exit Property
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, CStr(Pfx)) Then StrPfx = Pfx: Exit Property
Next
End Property

Property Get SubStrCnt&(A, SubStr)
Dim P&: P = 1
Dim L%: L = Len(SubStr)
Dim O%
While P > 0
    P = InStr(P, A, SubStr)
    If P = 0 Then SubStrCnt = O: Exit Property
    O = O + 1
    P = P + L
Wend
SubStrCnt = O
End Property

Property Get ToStr$(A)
If IsPrim(A) Then ToStr = A: Exit Property
If IsNothing(A) Then ToStr = "#Nothing": Exit Property
If IsEmpty(A) Then ToStr = "#Empty": Exit Property
If IsObject(A) Then
    Dim T$
    T = TypeName(A)
    Select Case T
    Case "CodeModule"
        Dim M As CodeModule
        Set M = A
        ToStr = FmtQQ("*Md{?}", M.Parent.Name)
        Exit Property
    End Select
    ToStr = "*" & T
    Exit Property
End If

If IsArray(A) Then
    Dim Ay: Ay = A: ReDim Ay(0)
    T = TypeName(Ay(0))
    ToStr = "*[" & T & "]"
    Exit Property
End If
Stop
End Property

Property Get TrimWhite$(A)
TrimWhite = RTrimWhite(LTrimWhite(A))
End Property

Sub StrBrw(A, Optional Fnn$)
Dim T$: T = TmpFt("StrBrw", Fnn$)
StrWrt A, T
FtBrw T
End Sub

Sub StrWrt(A, Ft)
Fso.CreateTextFile(Ft, True).Write A
End Sub

Sub ZZ__Tst()
ZZ_InstrN
ZZ_SubStrCnt
End Sub

Private Sub ZZ_InstrN()
Dim Act&, Exp&, S, SubStr, N%

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 1
Exp = 1
Act = InstrN(S, SubStr, N)
Stop
'Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 2
Exp = 6
Act = InstrN(S, SubStr, N)
Stop
'Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 3
Exp = 11
Act = InstrN(S, SubStr, N)
Stop
'Ass Exp = Act

'    12345678901234
S = ".aaaa.aaaa.bbb"
SubStr = "."
N = 4
Exp = 0
Act = InstrN(S, SubStr, N)
Stop
'Ass Exp = Act
End Sub

Private Sub ZZ_SubStrCnt()
Ass SubStrCnt("aaaa", "aa") = 2
Ass SubStrCnt("aaaa", "a") = 4
End Sub
