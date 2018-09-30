VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WhMd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Nm As WhNm
Dim X_InCmpTy() As vbext_ComponentType
Property Get InCmpTy() As vbext_ComponentType()
InCmpTy = X_InCmpTy
End Property
Property Let InCmpTy(V() As vbext_ComponentType)
X_InCmpTy = V
End Property

