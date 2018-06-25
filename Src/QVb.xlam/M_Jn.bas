Attribute VB_Name = "M_Jn"
Option Explicit

Property Get JnComma$(Ay)
JnComma = Join(AySy(Ay), ",")
End Property

Property Get JnCommaSpc(Ay)
JnCommaSpc = Join(AySy(Ay), ", ")
End Property

Property Get JnCrLf$(Ay, Optional WithIx As Boolean)
If WithIx Then
    Dim O$(), J%
    For J = 0 To UB(Ay)
        Push O, J & ": " & Ay(J)
    Next
    JnCrLf = Join(O, vbCrLf)
Else
    JnCrLf = Join(AySy(Ay), vbCrLf)
End If
End Property

Property Get JnDblCrLf$(Ay)
JnDblCrLf = Join(AySy(Ay), vbCrLf & vbCrLf)
End Property

Property Get JnQDblComma$(Ay)
JnQDblComma = JnComma(AyQuoteDbl(AySy(Ay)))
End Property

Property Get JnQDblSpc$(Ay)
JnQDblSpc = JnSpc(AyQuoteDbl(AySy(Ay)))
End Property

Property Get JnQSngComma$(Ay)
JnQSngComma = JnComma(AyQuoteSng(AySy(Ay)))
End Property

Property Get JnQSngSpc$(Ay)
JnQSngSpc = JnSpc(AyQuoteSng(AySy(Ay)))
End Property

Property Get JnQSqBktComma$(Ay)
JnQSqBktComma = JnComma(AyQuoteSqBkt(AySy(Ay)))
End Property

Property Get JnQSqBktSpc$(Ay)
JnQSqBktSpc = JnSpc(AyQuoteSqBkt(AySy(Ay)))
End Property

Property Get JnSpc$(Ay)
JnSpc = Join(AySy(Ay), " ")
End Property

Property Get JnTab$(Ay)
JnTab = Join(AySy(Ay), vbTab)
End Property

Property Get JnVBar$(Ay)
JnVBar = Join(AySy(Ay), "|")
End Property
