VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "P3NmRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A_Nm$
Private A_Er$()
Private A_Ok$()
Function Init(Nm$, Ok$(), Er$()) As P3NmRslt
A_Nm = Nm
A_Er = Er
A_Ok = Ok
Set Init = Me
End Function
Property Get Nm$()
Nm = A_Nm
End Property
Property Get Er() As String()
Er = A_Er
End Property
Property Get Ok() As String()
Ok = A_Ok
End Property

