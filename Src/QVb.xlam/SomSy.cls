VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomSy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Sy$()
Public Som As Boolean

Property Let Sy(V$())
B_Sy = V
End Property

Property Get Sy() As String()
Sy = B_Sy
End Property

Friend Property Get Init(Sy$()) As SomSy
B_Sy = Sy
Som = True
Set Init = Me
End Property
