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
Private A_LnxAy() As Lnx
Public BlkTyStr$
Friend Property Get Init(Ay() As Lnx, BlkTyStr$) As Blk
A_Lnxs = Ay
Me.BlkTyStr = BlkTyStr
Set Init = Me
End Property
Property Get LnxAy() As Lnx()
LnxAy = A_LnxAy
End Property
