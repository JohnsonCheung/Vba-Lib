Attribute VB_Name = "M_Macro"
Option Explicit

Function MacroNy(MacroStr, Optional ExclBkt As Boolean, Optional Bkt$ = "{}") As String()
Dim Q1$, Q2$
With BrkQuote(Bkt)
    Q1 = .S1
    Q2 = .S2
End With
End Function


