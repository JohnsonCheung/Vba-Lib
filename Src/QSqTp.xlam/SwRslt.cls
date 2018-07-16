VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SwRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Sw As New Sw
Private B_ErLy$()
Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("SwRslt", "Sw ErLy", Sw, SyObj(ErLy))
End Property
Property Get ErLy() As String()
ErLy = B_ErLy
End Property
Property Let ErLy(V$())
B_ErLy = V
End Property
