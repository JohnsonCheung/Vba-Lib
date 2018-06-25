VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCGp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_LBCAy() As LBC
Public A$, IsVF As Boolean

Property Get LBCAy() As LBC()
LBCAy = A_LBCAy
End Property

Property Get LBCVy()
LBCVy = A_LBCAy
End Property

Friend Property Get Init(A$, IsVF As Boolean) As LABCGp
Me.A = A
Me.IsVF = IsVF
Set Init = Me
End Property

Sub PushLxBC(Lx%, B$, C$)
Stop '
'PushObj A_LBCAy, LBC(Lx, B, C)
End Sub
