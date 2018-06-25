VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public K$, V

Property Get ToStr$()
ToStr = FmtQQ("KeyVal(K(?) V(?))", K, V)
End Property

Friend Property Get Init(K, V) As KeyVal
Me.K = K
Asg V, Me.V
Set Init = Me
End Property
