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
Public Sw As New Sw, Er As New Er
Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("SwRslt", "Sw Er", Sw, Er)
End Property
