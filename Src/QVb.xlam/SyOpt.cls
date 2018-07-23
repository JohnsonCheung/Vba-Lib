VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Sy$()
Public Som As Boolean

Property Let Sy(V$())
A_Sy = V
End Property

Friend Property Get Init(Sy$()) As SyOpt
A_Sy = Sy
Som = True
Set Init = Me
End Property

Property Get Sy() As String()
Sy = A_Sy
End Property
