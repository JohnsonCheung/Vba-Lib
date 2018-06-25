VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NmRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Nm$
Private A_ErLy$()

Property Get ErLy() As String()
ErLy = A_ErLy
End Property

Property Get ToStr$()
ToStr = Tag_NyStr_ObjAp("NmRslt", "Nm ErLy", StrObj(Nm), SyObj(A_ErLy))
End Property

Friend Property Get Init(Nm$, ErLy$()) As NmRslt
Me.Nm = Nm
A_ErLy = ErLy
Set Init = Me
End Property
