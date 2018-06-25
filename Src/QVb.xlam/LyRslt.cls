VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LyRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_Ly$()
Private A_ErLy$()

Property Get ErLy() As String()
ErLy = A_ErLy
End Property

Property Get Ly()
Ly = A_Ly
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("LyRslt", "Ly ErLy", SyObj(A_Ly), SyObj(A_ErLy))
End Property

Friend Property Get Init(Ly$(), ErLy$()) As LyRslt
A_Ly = Ly
A_ErLy = ErLy
Set Init = Me
End Property
