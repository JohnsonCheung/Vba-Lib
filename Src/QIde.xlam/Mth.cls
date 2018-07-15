VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Md As CodeModule, Nm$

Friend Property Get Init(Md As CodeModule, Nm$) As Mth
Set Me.Md = Md
Me.Nm = Nm
Set Init = Me
End Property
