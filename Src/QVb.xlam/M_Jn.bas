Attribute VB_Name = "M_Jn"
Option Explicit

Function JnComma$(Ay)
JnComma = Join(AySy(Ay), ",")
End Function

Function JnCommaSpc(Ay)
JnCommaSpc = Join(AySy(Ay), ", ")
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
JnDblCrLf = Join(AySy(Ay), vbCrLf & vbCrLf)
End Function

Function JnQDblComma$(Ay)
JnQDblComma = JnComma(AyQuoteDbl(AySy(Ay)))
End Function

Function JnQDblSpc$(Ay)
JnQDblSpc = JnSpc(AyQuoteDbl(AySy(Ay)))
End Function

Function JnQSngComma$(Ay)
JnQSngComma = JnComma(AyQuoteSng(AySy(Ay)))
End Function

Function JnQSngSpc$(Ay)
JnQSngSpc = JnSpc(AyQuoteSng(AySy(Ay)))
End Function

Function JnQSqBktComma$(Ay)
JnQSqBktComma = JnComma(AyQuoteSqBkt(AySy(Ay)))
End Function

Function JnQSqBktSpc$(Ay)
JnQSqBktSpc = JnSpc(AyQuoteSqBkt(AySy(Ay)))
End Function

Function JnSpc$(Ay)
JnSpc = Join(AySy(Ay), " ")
End Function

Function JnTab$(Ay)
JnTab = Join(AySy(Ay), vbTab)
End Function

Function JnVBar$(Ay)
JnVBar = Join(AySy(Ay), "|")
End Function
