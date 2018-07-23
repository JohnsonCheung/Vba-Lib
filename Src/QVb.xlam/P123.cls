VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public P1$, P2$, P3$

Friend Property Get Init(P1, P2, P3) As P123
Me.P1 = P1
Me.P2 = P2
Me.P3 = P3
Set Init = Me
End Property

Property Get ToStr$()
ToStr = FmtQQ("P123(|P1(?)|P2(?)|P3(?)|P123)", P1, P2, P3)
End Property
