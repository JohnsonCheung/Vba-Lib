VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoolOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Bool As Boolean, Som As Boolean

Friend Property Get Init(Bool As Boolean) As BoolOpt
Me.Bool = Bool
Me.Som = True
Set Init = Me
End Property

Property Get ToStr$()
If Som Then
    ToStr = FmtQQ("BoolOpt(?)", Bool)
Else
    ToStr = "BoolOpt(None)"
End If
End Property


