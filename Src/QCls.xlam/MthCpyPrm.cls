VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MthCpyPrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public SrcMth As Mth, ToMd As CodeModule
Friend Property Get Init(SrcMth As Mth, ToMd As CodeModule) As MthCpyPrm
Set Me.SrcMth = SrcMth
Set Me.ToMd = ToMd
End Property
