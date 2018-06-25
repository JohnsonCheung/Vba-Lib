VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCAyRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_LABCAy() As LABC
Private A_ErLy$()

Property Get Init(A() As LABC, ErLy$()) As LABCAyRslt
A_LABCAy = A
A_ErLy = ErLy
Set Init = Me
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("LABCAyRslt", "LABCAy ErLy", LABCAy_ToStr, SyObj(A_ErLy))
End Property

Private Property Get LABCAy_ToStr$()
End Property
