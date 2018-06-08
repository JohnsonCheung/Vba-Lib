VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P123"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public P1$, P2$, P3$
Property Get ToStr$()
ToStr = FmtQQ("P123(P1(?) P2(?) P3(?))", P1, P2, P3)
End Property
