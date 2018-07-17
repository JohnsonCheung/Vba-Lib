VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ValOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Som As Boolean, V

Friend Property Get Init(V) As ValOpt
Asg V, Me.V
Som = True
Set Init = Me
End Property
