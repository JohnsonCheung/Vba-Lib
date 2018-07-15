VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LnxsRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_LnxAy() As Lnx
Private A_ErLy$()

Property Get ErLy() As String()
ErLy = A_ErLy
End Property

Property Get LnxAy() As Lnx()
LnxAy = A_LnxAy
End Property

Friend Property Get Init(LnxAy() As Lnx, ErLy$()) As LnxsRslt
A_LnxAy = LnxAy
A_ErLy = ErLy
Set Init = Me
End Property
