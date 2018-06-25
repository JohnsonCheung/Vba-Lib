VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoFmtrRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Ly$()
Private A_ErLy$()
Public WsFmtr As WsFmtr
Friend Property Get Init(Ly$(), ErLy$(), WsFmtr As WsFmtr) As LoFmtrRslt
A_Ly = Ly
Set A_ErLy = ErLy
Set A_WsFmtr = WsFmtr
Set Init = Me
End Property

Property Get Ly() As String()
Ly = A_Ly
End Property

Property Get ErLy() As String()
ErLy = A_ErLy
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("LoFmtrRslt", "Ly WsFmtr Er", SyObj(A_Ly), A_WsFmtr, SyObj(A_ErLy))
End Property
