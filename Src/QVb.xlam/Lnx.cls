VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Lnx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Lx%, A_Lin$

Property Get Lin$()
Lin = A_Lin
End Property

Property Get Lx%()
Lx = A_Lx
End Property

Property Get ToStr$()
ToStr = FmtQQ("Lnx(?: ?", A_Lx, A_Lin)
End Property

Friend Property Get Init(Lin, Lx%) As Lnx
A_Lx = Lx
A_Lin = Lin
Set Init = Me
End Property
