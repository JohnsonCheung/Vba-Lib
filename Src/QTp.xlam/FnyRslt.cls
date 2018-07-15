VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FnyRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Fny$()
Private A_ErLy

Property Get ErLy() As String()
ErLy = A_ErLy
End Property

Property Get Fny() As String()
Fny = A_Fny
End Property

Property Get Init(Fny$(), ErLy$()) As FnyRslt
A_Fny = Fny
A_ErLy = ErLy
Set Init = Me
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("FnyRslt", "Fny ErLy", SyObj(Fny), SyObj(ErLy))
End Property
