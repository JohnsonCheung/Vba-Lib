VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Mth"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Md As CodeModule
Public Nm$
Property Get Init(A As CodeModule, MthNm) As Mth
Set Md = A
Nm = MthNm
Set Init = Me
End Property
Property Get ToStr$()
ToStr = "Mth(" & ZMd_DNm(Md) & "." & Nm & ")"
End Property
