VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Ds"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_DtAy() As Dt, DsNm$
Property Get DtAy() As Dt()
DtAy = A_DtAy
End Property
Property Get Init(A() As Dt, Optional DsNm$ = "Ds") As Ds
A_DtAy = A
A_DsNm = DsNm
Set Init = Me
End Property
