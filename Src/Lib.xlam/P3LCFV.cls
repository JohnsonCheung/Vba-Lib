VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3LCFV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Lx%
Public Cno%
Public Fld$
Public Val
Property Get Dr() As Variant()
Dr = Array(Lx, Cno, Fld, Val)
End Property
Property Get Show() As String()
Show = ApSy("P3LCFV(? ? ? ?)", Lx, Cno, Fld, Val)
End Property
