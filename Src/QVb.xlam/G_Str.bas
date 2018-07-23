Attribute VB_Name = "G_Str"
Option Explicit

Function FstChr$(A)
FstChr = Left(A, 1)
End Function

Function InstrN&(S, SubStr, N%)
Dim P&, J%
For J = 1 To N
    P = InStr(P + 1, S, SubStr)
    If P = 0 Then Exit Function
Next
InstrN = P
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

Function Quote$(A, QuoteStr$)
With BrkQuote(QuoteStr)
    Quote = .S1 & A & .S2
End With
End Function

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
