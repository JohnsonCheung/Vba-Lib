VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MthOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Som As Boolean
Public Mth As Mth

Friend Property Get Init(A As Mth) As MthOpt
Som = True
Set Mth = A
Set Init = Me
End Property

Property Get ToStr$()
If Som Then
    ToStr = FmtQQ("MthOpt(?)", Mth.ToStr)
Else
    ToStr = FmtQQ("MthOpt(None)")
End If
End Property

