VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "StrObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$
Friend Property Get Init(Str) As StrObj
A = Str
Set Init = Me
End Property

Property Get ToStr$()
ToStr = Tag("StrObj", A)
End Property
