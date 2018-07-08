Attribute VB_Name = "M_Brk"
Option Explicit

Property Get Brk(A, Sep, Optional NoTrim As Boolean) As S1S2
Const CSub$ = "Brk"
Dim P&: P = InStr(A, Sep)
If P = 0 Then
    Er CSub, "{S} does not contains {Sep}", A, Sep
End If
Set Brk = BrkAt(A, P, Len(Sep), NoTrim)
End Property

Property Get Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk1 = Brk1__(A, P, Sep, NoTrim)
End Property

Property Get Brk1Rev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
Set Brk1Rev = Brk1__(A, P, Sep, NoTrim)
End Property

Property Get Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S2 = A
    Else
        O.S2 = Trim(A)
    End If
    Brk2 = O
    Exit Property
End If
Set Brk2 = BrkAt(A, P, Len(Sep), NoTrim)
End Property

Property Get BrkBkt(A, Optional Bkt$ = "()") As P123
Const CSub$ = "BktPos"
Dim P As FmToPos: Set P = BrkBktPos(A, Bkt)
Dim L1%, P2%, L2%, P3% 'L for Len, P for Position
    Dim Q1$, Q2$
        BrkQUote_Asg Bkt, Q1, Q2
    L1 = P.FmPos - 1
    P2 = L1 + Len(Q1) + 1
    L2 = P.ToPos - P2
    P3 = P.ToPos + Len(Q2)
Dim A1$, A2$, A3$
A1 = Left(A, L1)
A2 = Mid(A, P2, L2)
A3 = Mid(A, P3)
Set BrkBkt = P123(A1, A2, A3)
End Property

Property Get BrkBktPos(A, Optional Bkt$ = "()") As FmToPos
Const CSub$ = "BrkBktPos"
Dim Q1$, Q2$
    BrkQUote_Asg Bkt, Q1, Q2
Dim IsBkt As Boolean
    Select Case True
    Case Bkt = "()", Bkt = "[]", Bkt = "{}": IsBkt = True
    End Select
Dim P1%
    P1 = InStr(A, Q1)

If Not IsBkt Then
    Set BrkBktPos = FmToPos(P1, InStr(P1, A, Q2))
    Exit Property
End If

Dim P2%
    Dim NOpn%, J%
    For J = P1 + 1 To Len(A)
        Select Case Mid(A, J, 1)
        Case Q2
            If NOpn = 0 Then
                P2 = J
                Exit For
            End If
            NOpn = NOpn - 1
        Case Q1
            NOpn = NOpn + 1
        End Select
    Next
    If P2 = 0 Then
        Er CSub, "{Str} has {Q1} and {Q2} is not in pair", A, Q1, Q2
    End If
Set BrkBktPos = FmToPos(P1, P2)
End Property

Property Get BrkBoth(A, Sep, Optional NoTrim As Boolean) As S1S2
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
    Exit Property
End If
Set BrkBoth = BrkAt(A, P, Len(Sep), NoTrim)
End Property

Property Get BrkQuote(QuoteStr$) As S1S2
Dim L%: L = Len(QuoteStr)
Dim O As New S1S2
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
        Set O = Brk(QuoteStr, "*", NoTrim:=True)
    End If
End Select
Set BrkQuote = O
End Property

Property Get BrkRev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & A & "] does not contains Sep[" & Sep & "]"
BrkRev = BrkAt(A, P, Len(Sep), NoTrim)
End Property

Sub BrkAsg(A, Sep, O1$, O2$, Optional NoTrim As Boolean)
Brk(A, Sep, NoTrim).Asg O1, O2
End Sub

Sub BrkQUote_Asg(QuoteStr$, O1$, O2$)
With BrkQuote(QuoteStr)
    O1 = .S1
    O2 = .S2
End With
End Sub

Private Property Get Brk1__(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    Dim O As New S1S2
    If NoTrim Then
        O.S1 = A
    Else
        O.S1 = Trim(A)
    End If
    Set Brk1__ = O
    Exit Property
End If
Set Brk1__ = BrkAt(A, P, Len(Sep), NoTrim)
End Property

Sub ZZ__Tst()
ZZ_Brk1Rev
End Sub

Private Property Get BrkAt(A, P&, SepLen%, Optional NoTrim As Boolean) As S1S2
Dim O As New S1S2
With O
    If NoTrim Then
        .S1 = Left(A, P - 1)
        .S2 = Mid(A, P + SepLen)
    Else
        .S1 = Trim(Left(A, P - 1))
        .S2 = Trim(Mid(A, P + SepLen))
    End If
End With
Set BrkAt = O
End Property

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
