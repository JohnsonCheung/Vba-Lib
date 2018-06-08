Attribute VB_Name = "M_Brk"
Option Explicit

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
If P = 0 Then
    Dim O As S1S2
    If NoTrim Then
        O.S2 = A
    Else
        O.S2 = Trim(A)
    End If
    Brk2 = O
    Exit Function
End If
Brk2 = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Function Brk(A, Sep, Optional NoTrim As Boolean) As S1S2
Const CSub$ = "Brk"
Dim P&: P = InStr(A, Sep)
If P = 0 Then
    'Er CSub, "{S} does not contains {Sep}", A, Sep
    Stop
End If
Set Brk = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Sub BrkAsg(A, Sep, O1$, O2$, Optional NoTrim As Boolean)
Brk(A, Sep, NoTrim).Asg O1, O2
End Sub

Function BrkAt(A, P&, SepLen%, Optional NoTrim As Boolean) As S1S2
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
Set BrkBoth = BrkAt(A, P, Len(Sep), NoTrim)
End Function

Function BrkQuote(QuoteStr$) As S1S2
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
End Function
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


Function BrkRev(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStrRev(A, Sep)
If P = 0 Then Err.Raise "BrkRev: Str[" & A & "] does not contains Sep[" & Sep & "]"
BrkRev = BrkAt(A, P, Len(Sep), NoTrim)
End Function
Function BrkBktPos(A, Optional Bkt$ = "()") As FmToPos
Stop
End Function

Function BrkBkt(A, Optional Bkt$ = "()") As P123
Const CSub$ = "BktPos"
Dim Q1$, Q2$
    BrkQuote(Bkt).Asg Q1, Q2
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
    If ToPos = 0 Then
        Stop
        'Er CSub, "{Str} has {Q1} and {Q2} is not in pair", A, Q1, Q2
    End If
Dim L1%, N2%, L2%, N3%
Stop
Dim O As New P123
With O
    .P1 = Left(A, L1)
    .P2 = Mid(A, N2, L2)
    .P3 = Mid(A, N3)
End With
Set BrkBkt = O
End Function

Function Quote$(A, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & A & .S2
End With
End Function
