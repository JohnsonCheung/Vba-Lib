VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCItm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public A$, IsVF As Boolean
Private A_LBC() As LBC

Property Get LBCAy() As LBC
LBCAy = A_LBC
End Property

Property Get ToStr$()
Stop '
'ToStr = Tag("LABCItm", "A IsVF LBCAy", StrObj(A), BoolObj(IsVF), AyObj(A_LBC))
End Property

Friend Property Get Init(A$, IsVF As Boolean, LBCAy() As LBC) As LABCItm
Stop '
'Me.A = A
'Me.IsVF = IsVF
'A_LBCAy = LBCAy
End Property
