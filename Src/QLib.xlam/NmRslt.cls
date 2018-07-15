VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NmRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Nm$, Er As New Er

Property Get Lin$()
Lin = FmtQQ("? ? ?", C1_Lo, C2_Lo_Nam, Nm)
End Property

Property Get ToStr$()
ToStr = "Nm(?) " & Er.ToStr
End Property

Friend Property Get Init(Nm$, Er As Er) As NmRslt
Me.Nm = Nm
Set Me.Er = Er
Set Init = Me
End Property
