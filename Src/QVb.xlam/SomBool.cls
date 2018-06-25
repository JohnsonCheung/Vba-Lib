VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomBool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Bool As Boolean, Som As Boolean

Property Get ToStr$()
If Som Then
    ToStr = FmtQQ("SomBool(?)", Bool)
Else
    ToStr = "SomBool(None)"
End If
End Property

Friend Property Get Init(Bool As Boolean) As SomBool
Me.Bool = Bool
Set Init = Me
End Property
