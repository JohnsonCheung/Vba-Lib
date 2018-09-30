VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Blk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private X_LnxAy() As Lnx
Public BlkTyStr$
Property Get LnxAy() As Lnx()
LnxAy = X_LnxAy
End Property
Property Let LnxAy(V() As Lnx)
X_LnxAy = V
End Property
