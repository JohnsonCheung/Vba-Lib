Attribute VB_Name = "M_Str"
'Option Explicit
'Function ValStr$(A)
'If IsPrim(A) Then ValStr = A: Exit Function
'If IsNothing(A) Then ValStr = "#Nothing": Exit Function
'If IsEmpty(A) Then ValStr = "#Empty": Exit Function
'Dim T$
'If IsObject(A) Then
'    T = TypeName(A)
'    Select Case T
'    Case "CodeModule"
'        ValStr = FmtQQ("*Md{?}", CvMdx(A).Nm)
'        Exit Function
'    End Select
'    ValStr = FmtQQ("*?{?}", T, ObjS(A))
'    Exit Function
'End If
'If IsArray(A) Then
'    ValStr = "*Array"
'    Exit Function
'End If
'Stop
'End Function

Function AlignL$(A, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "AlignL"
If ErIfNotEnoughWdt And DoNotCut Then
    Stop
    'Er CSub, "Both {ErIfNotEnoughWdt} and {DontCut} cannot be True", ErIfNotEnoughWdt, DoNotCut
End If
Stop
Dim S$: 'S = ValStr(A)
AlignL = StrAlignL(S, W, ErIfNotEnoughWdt, DoNotCut)
End Function
Function StrAlignL$(S$, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "StrAlignL"
Dim L%: L = Len(S)
If L > W Then
    If ErIfNotEnoughWdt Then
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
Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
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
Function IsWhite(A) As Boolean
Dim B$: B = Left(A, 1)
IsWhite = True
If B = " " Then Exit Function
If B = vbCr Then Exit Function
If B = vbLf Then Exit Function
If B = vbTab Then Exit Function
IsWhite = False
End Function
Function LTrimWhite$(A)
Dim J%
    For J = 1 To Len(A)
        If Not IsWhite(Mid(A, J, 1)) Then Exit For
    Next
LTrimWhite = Left(A, J)
End Function

Function TrimWhite$(A)
TrimWhite = RTrimWhite(LTrimWhite(A))
End Function

Function Dft(Val, DftV)
If IsEmp(Val) Then
   Dft = DftV
Else
   Dft = Val
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
    Stop
    'Debug.Print ErMsgLines("Esc", "Warning: escaping a {Str} of {FmStrSub} to {ToSubStr} is found that {Str} contains some {ToSubStr}.  This will make the string chagned after UnEsc", A, Fm, ToStr)
End If
Esc = Replace(A, Fm, ToStr)
End Function

Function EscCr$(A)
EscCr = Esc(A, vbCr, "\r")
End Function

Function UnEscCr$(A)
UnEscCr = Replace(A, "\r", vbCr)
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

Function LasChr$(A)
LasChr = Right(A, 1)
End Function



Function Prepend$(S, P)
If S <> "" Then Prepend = P & S
End Function

Function PrependDash$(S)
PrependDash = Prepend(S, "-")
End Function




Function AyTrim(A) As String()
Stop
'If AyIsEmp(A) Then Exit Function
'Dim U&
'    U = UB(A)
'Dim O$()
'    Dim J&
'    ReDim O(U)
'    For J = 0 To U
'        O(J) = Trim(A(J))
'    Next
'AyTrim = O
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

Sub StrBrw(A, Optional Fnn$)
Stop
'Dim T$: T = TmpFt("StrBrw", Fnn$)
'StrWrt A, T
'FtBrw T
End Sub

Function StrDup$(N%, S)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function StrPfx$(A, PfxAy$())
If AyIsEmp(PfxAy) Then Exit Function
Dim Pfx
For Each Pfx In PfxAy
    Stop
'    If HasPfx(A, CStr(Pfx)) Then StrPfx = Pfx: Exit Function
Next
End Function

Sub StrWrt(A, Ft)
Stop
'Fso.CreateTextFile(Ft, True).Write A
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

Function TmpFfn(Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
If Fnn0 = "" Then
    Fnn = TmpNm
Else
    Fnn = Fnn0
End If
TmpFfn = TmpPth(Fdr) & Fnn & Ext
End Function

Function UnEscSpc$(A)
UnEscSpc = Replace(A, "~", " ")
End Function

Function UnEscTab(A)
UnEscTab = Replace(A, "\t", "~")
End Function


Private Sub InstrN__Tst()
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

Sub ZZ_SubStrCnt()
Ass SubStrCnt("aaaa", "aa") = 2
Ass SubStrCnt("aaaa", "a") = 4
End Sub
