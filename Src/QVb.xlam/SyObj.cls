VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Sy$()

Property Get ToStr$()
ToStr = Tag("SyObj", JnCrLf(B_Sy, WithIx:=True))
End Property

Friend Property Get Init(Sy$()) As SyObj
B_Sy = Sy
Set Init = Me
End Property
