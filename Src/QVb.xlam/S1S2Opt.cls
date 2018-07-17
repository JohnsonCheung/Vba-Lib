VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1S2Opt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public S1S2 As S1S2, Som As Boolean

Friend Property Get Init(S1S2 As S1S2) As S1S2Opt
Set Me.S1S2 = S1S2
Som = True
Set Init = Me
End Property
