VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomKeyVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public KeyVal As KeyVal
Public Som As Boolean

Property Get Init(A As KeyVal) As SomKeyVal
Set KeyVal = A
Som = True
Set Init = Me
End Property

Property Get ToStr$()
If Som Then
    ToStr = Tag_NyStr_ObjAp("SomKeyVal", "Som", KeyVal)
Else
    ToStr = "SomKeyVal(None)"
End If
End Property
