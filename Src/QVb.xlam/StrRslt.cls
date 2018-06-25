VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_ErLy$()
Public Str$

Property Get ErLy() As String()
ErLy = A_ErLy
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("StrRslt", "Str ErLy", StrObj(Str), SyObj(ErLy))
End Property

Friend Property Get Init(S, ErLy$()) As StrRslt
Me.Str = S
A_ErLy = ErLy
Set Init = Me
End Property
