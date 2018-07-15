VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LBC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Lx%, B$, C$

Property Get ToStr$()
Dim S$: S = FmtQQ("Lx(?) B(?) C(?)", Lx, B, C)
ToStr = Tag("LBC", S)
End Property

Friend Property Get Init(Lx%, B$, C$) As LBC
Me.Lx = Lx
Me.B = B
Me.C = C
Set Init = Me
End Property
