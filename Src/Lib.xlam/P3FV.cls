VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3FV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Bet As New LABCs
Public Fml As New LABCs
Public Lbl As New LABCs
Public Tit As New LABCs


Private Sub Class_Initialize()
Bet.InitByT1 C1_Bet
Fml.InitByT1 C1_Fml
Lbl.InitByT1 C1_Lbl
Tit.InitByT1 C1_Tit
End Sub
