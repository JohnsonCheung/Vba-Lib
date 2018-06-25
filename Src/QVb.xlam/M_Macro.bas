Attribute VB_Name = "M_Macro"
Option Explicit

Property Get MacroNy(MacroStr, Optional ExclBkt As Boolean, Optional Bkt$ = "{}") As String()
Dim Q1$, Q2$
With BrkQuote(Bkt)
    Q1 = .S1
    Q2 = .S2
End With
If Q1 = Q2 Then Stop
If Len(Q1) <> 1 Then Stop
If Len(Q2) <> 1 Then Stop

Dim A$(): A = Split(MacroStr, Q1)
Dim O$(), J%
For J = 1 To UB(A)
    Push O, TakBef(A(J), Q2)
Next
If Not ExclBkt Then
    O = AyAddPfxSfx(O, Q1, Q2)
End If
MacroNy = O
End Property
