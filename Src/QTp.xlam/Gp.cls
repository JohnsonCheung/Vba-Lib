VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Gp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_LnxAy() As Lnx

Property Get LnxAy() As Lnx()
LnxAy = A_LnxAy
End Property

Property Get ToStr$()
ToStr = Tag("Gp", OyToStr(A_LnxAy))
End Property

Friend Property Get Init(A() As Lnx) As Gp
A_LnxAy = A
Set Init = Me
End Property
