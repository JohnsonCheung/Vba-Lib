VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CnoVal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Cno%, Val

Friend Property Get Init(Cno%, Val) As CnoVal
Me.Cno = Cno
Asg Val, Me.Val
Set Init = Me
End Property

Property Get ToStr$()
ToStr = FmtQQ("CnoVal(Cno(?) Val(?) V-Ty(?))", Cno, Val, TypeName(Val))
End Property
