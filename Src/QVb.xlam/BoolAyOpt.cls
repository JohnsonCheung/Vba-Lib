VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BoolAyOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_BoolAy() As Boolean, Som As Boolean

Friend Property Get Init(A() As Boolean) As BoolAyOpt
A_BoolAy = A
Som = True
Set Init = Me
End Property
