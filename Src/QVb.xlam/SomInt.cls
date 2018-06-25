VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public I%, Som As Boolean

Friend Property Get Init(I%) As SomInt
Me.I = I
Som = True
Set Init = Me
End Property
