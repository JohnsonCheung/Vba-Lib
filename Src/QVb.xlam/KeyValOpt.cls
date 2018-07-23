VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "KeyValOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public KeyVal As KeyVal
Public Som As Boolean

Property Get Init(A As KeyVal) As KeyValOpt
Set KeyVal = A
Som = True
Set Init = Me
End Property

Property Get ToStr$()
If Som Then
    ToStr = Tag_NyStr_ObjAp("KeyValOpt", "Som", KeyVal)
Else
    ToStr = "KeyValOpt(None)"
End If
End Property


