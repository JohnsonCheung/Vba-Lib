VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Macro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Macro$
Property Get Ny(Optional ExclBkt As Boolean, Optional Bkt$ = "{}") As String()
Dim Q1$, Q2$
With BrkQuote(Bkt)
    Q1 = .S1
    Q2 = .S2
End With
If Q1 = Q2 Then Stop
If Len(Q1) <> 1 Then Stop
If Len(Q2) <> 1 Then Stop

Dim Ay$(): Ay = Split(Macro, Q1)
Dim O$(), J%
For J = 1 To UB(Ay)
    Push O, TakBef(Ay(J), Q2)
Next
If Not ExclBkt Then
    O = AyAddPfxSfx(O, Q1, Q2)
End If
Ny = O
End Property
