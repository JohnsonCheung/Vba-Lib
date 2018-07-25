Attribute VB_Name = "M_Brk"
Option Explicit

Function Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk1 = Brk1__X(A, P, Sep, NoTrim)
End Function

Function BrkAt(A, P&, Sep, Optional NoTrim As Boolean) As S1S2
Dim SepLen%: SepLen = Len(Sep)
Dim S1$, S2$
If NoTrim Then
    S1 = Left(A, P - 1)
    S2 = Mid(A, P + SepLen)
Else
    S1 = Trim(Left(A, P - 1))
    S2 = Trim(Mid(A, P + SepLen))
End If
Set BrkAt = S1S2(S1, S2)
End Function

Function Brk(A, Sep, Optional NoTrim As Boolean) As S1S2
Const CSub$ = "Brk"
Dim P&: P = InStr(A, Sep)
If P = 0 Then
    Er CSub, "{S} does not contains {Sep}", A, Sep
End If
Set Brk = BrkAt(A, P, Sep, NoTrim)
End Function

Function Brk1Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
Set Brk1Rev = Brk1__X(A, P, Sep, NoTrim)
End Function

Function Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk2 = Brk2__X(A, P, Sep, NoTrim)
End Function

Function Brk2Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
Set Brk2Rev = Brk2__X(A, P, Sep, NoTrim)
End Function

Function BrkBkt(A, Optional Bkt$ = "()") As P123
Const CSub$ = "BktPos"
Dim P As FmToPos: Set P = BrkBktPos(A, Bkt)
Dim L1%, P2%, L2%, P3% 'L for Len, P for Position
    Dim Q1$, Q2$
        BrkQuoteAsg Bkt, Q1, Q2
    L1 = P.FmPos - 1
    P2 = L1 + Len(Q1) + 1
    L2 = P.ToPos - P2
    P3 = P.ToPos + Len(Q2)
Dim A1$, A2$, A3$
A1 = Left(A, L1)
A2 = Mid(A, P2, L2)
A3 = Mid(A, P3)
Set BrkBkt = P123(A1, A2, A3)
End Function

Function BrkBktPos(A, Optional Bkt$ = "()") As FmToPos
Const CSub$ = "BrkBktPos"
Dim Q1$, Q2$
    BrkQuoteAsg Bkt, Q1, Q2
Dim IsBkt As Boolean
    Select Case True
    Case Bkt = "()", Bkt = "[]", Bkt = "{}": IsBkt = True
    End Select
Dim P1%
    P1 = InStr(A, Q1)

If Not IsBkt Then
    Set BrkBktPos = FmToPos(P1, InStr(P1, A, Q2))
    Exit Function
End If
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
End Function

Sub BrkQuoteAsg(QuoteStr$, O1$, O2$)
S1S2_Asg BrkQuote(QuoteStr), O1, O2
End Sub

Function BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim S1$, S2$
Select Case L
Case 0:
Case 1
    S1 = QuoteStr
    S2 = QuoteStr
Case 2
    S1 = Left(QuoteStr, 1)
    S2 = Right(QuoteStr, 1)
Case Else
    If InStr(QuoteStr, "*") > 0 Then
        Set BrkQuote = Brk(QuoteStr, "*", NoTrim:=True)
        Exit Function
    End If
    Stop
End Select
Set BrkQuote = S1S2(S1, S2)
End Function

Function BrkRev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & A & "] does not contains Sep[" & Sep & "]"
BrkRev = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Function Brk1__X(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk1__X = S1S2("", A)
    Else
        Set Brk1__X = S1S2("", Trim(A))
    End If
    Exit Function
End If
End Function

Function Brk2__X(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk2__X = S1S2("", A)
    Else
        Set Brk2__X = S1S2("", Trim(A))
    End If
    Exit Function
End If
End Function

Sub ZZZ__Tst()
ZZ_Brk1Rev
End Sub

Private Sub ZZ_Brk1Rev()
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

Private Sub ZZ_BrkBkt()
Dim A$, Act$
A = "aa<xx>bbb</xx>cccc"
Act = BrkBkt(A, "<xx>*</xx>").ToStr
Ass Act = RplVBar("P123(|P1(aa)|P2(bbb)|P3(cccc)|P123)")

A = "aaaa((a),(b))xxx"
Act = BrkBkt(A).ToStr
Ass Act = RplVBar("P123(|P1(aaaa)|P2((a),(b))|P3(xxx)|P123)")
End Sub
