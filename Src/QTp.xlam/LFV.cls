VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LFV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Lx%, F$, V$

Friend Property Get Init(Lx%, F$, V$) As LFV
Me.Lx = Lx
Me.F = F
Me.V = V
Set Init = Me
End Property
