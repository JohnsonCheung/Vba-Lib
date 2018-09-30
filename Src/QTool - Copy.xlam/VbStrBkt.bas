Attribute VB_Name = "VbStrBkt"
Option Explicit
Sub Z_BktPos()
Dim A$
A = "(A(B)A)A"
Ept = ApIntAy(1, 7)
GoSub Tst
Exit Sub
Tst:
    Act = BktPos(A)
    C
    Return
End Sub

Function BktPos(A) As Integer()
Dim P%, EndP%
P = InStr(A, "("): If P = 0 Then Exit Function
EndP = BktEndPos(A, P)
If EndP > P Then BktPos = ApIntAy(P, EndP)
End Function

Function ShfBktStr$(OLin$)
Dim O$
O = BktStr(OLin): If O = "" Then Exit Function
ShfBktStr = O
OLin = TakAftBkt(OLin)
End Function

Function SqBktEndPos%(A, Optional FmPos% = 1)
SqBktEndPos = BktXEndPos(A, "[", "]", FmPos)
End Function

Function BktEndPos%(A, Optional FmPos% = 1)
BktEndPos = BktXEndPos(A, "(", ")", FmPos)
End Function
Private Sub Z_BktXEndPos()
Dim A$, Q1$, Q2$, FmPos%
Q1 = "[": Q2 = "]": FmPos = 1
Ept = 7: A = "[ksdf]]]dkf": GoSub Tst
Ept = 0: A = "[[]":         GoSub Tst
Exit Sub
Tst:
    Act = BktXEndPos(A, Q1, Q2, FmPos)
    Return
End Sub
Function BktXEndPos%(A, Q1$, Q2$, Optional FmPos% = 1)
If Mid(A, FmPos, 1) <> Q1 Then Stop: Exit Function
Dim L$, Fm%, N As Byte, J%
L = A
Fm = FmPos + 1
GoSub R
Exit Function
R:
    J = J + 1: If J > 1000 Then Stop
    If N < 0 Then Stop
    If Fm > Len(L) Then Exit Function
    Dim P%, IsO As Boolean: GoSub P_IsO
    If P = 0 Then Exit Function
    Fm = P + 1
    If IsO Then
        N = N + 1
        GoTo R
    End If
    If N = 0 Then
        BktXEndPos = P
        Exit Function
    End If
    N = N - 1
    GoTo R
P_IsO:
    Dim C%
    IsO = False
    P = InStr(Fm, L, Q1): If P = 0 Then P = InStr(Fm, L, Q2): Return
    C = InStr(Fm, L, Q2): If C = 0 Then Return
    If C < P Then P = C: Return
    IsO = True
    Return
End Function

Private Sub Z_TakAftBkt()
Dim A$
A = "(lsk(aa)df lsdkfj) A"
Ept = " A"
GoSub Tst
Exit Sub
Tst:
    Act = TakAftBkt(A)
    C
    Return
End Sub

Function TakAftBkt$(A)
Dim P%(): P = BktPos(A)
If Sz(P) = 0 Then Exit Function
TakAftBkt = Mid(A, P(1) + 1)
End Function
Private Sub Z_ShfBktStr()
Dim A$, Ept1$
A$ = "(O$()) As X": Ept = "O$()": Ept1 = " As X": GoSub Tst
Exit Sub
Tst:
    Act = ShfBktStr(A)
    C
    Ass A = Ept1
    Return
End Sub
Sub Z_BktStr()
Dim A$
Ept = "AA":     A = "XXX(AA)XX":   GoSub Tst
Ept = "A$()A": A = "(A$()A)XX":   GoSub Tst
Ept = "O$()":   A = "(O$()) As X": GoSub Tst

Exit Sub
Tst:
    Act = BktStr(A)
    C
    Return
End Sub
Function BktStr$(A)
Dim P%()
P = BktPos(A)
If Sz(P) = 0 Then Exit Function
BktStr = Mid(A, P(0) + 1, P(1) - P(0) - 1)
End Function
