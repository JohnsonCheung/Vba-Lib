VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Macro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Public Macro$
'
'Friend Property Get Init(MacroStr) As Macro
'Macro = MacroStr
'Set Init = Me
'End Property
'
'Property Get Ny(Optional ExclBkt As Boolean, Optional Bkt$ = "{}") As String()
'Dim Q1$, Q2$
'With BrkQuote(Bkt)
'    Q1 = .S1
'    Q2 = .S2
'End With
'If Q1 = Q2 Then Stop
'If Len(Q1) <> 1 Then Stop
'If Len(Q2) <> 1 Then Stop
'
'Dim A$(): A = Split(Macro, Q1)
'Dim O$(), J%
'For J = 1 To UB(A)
'    Push O, TakBef(A(J), Q2)
'Next
'If Not ExclBkt Then
'    O = Ay(O).AddPfxSfx(Q1, Q2)
'End If
'Ny = O
'End Property
