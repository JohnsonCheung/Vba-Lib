VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PmRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Pm As Dictionary
Private A_ErLy$()
Friend Property Get Init(Pm As Dictionary, ErLy$()) As PmRslt
Set Me.Pm = Pm
A_ErLy = ErLy
Set Init = Me
End Property
Property Get ErLy() As String()
ErLy = A_ErLy
End Property

