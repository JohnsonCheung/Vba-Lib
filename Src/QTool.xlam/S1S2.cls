VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1S2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public S1$, S2$

Friend Property Get Init(S1, S2) As S1S2
Me.S1 = S1
Me.S2 = S2
Set Init = Me
End Property

Property Get ToStr$()
ToStr = FmtQQ("S1S2(S1(?) S2(?))", StrLin(S1), StrLin(S2))
End Property
