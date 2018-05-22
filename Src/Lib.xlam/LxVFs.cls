VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LxVFs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A_Ay() As LxVF
Function Init(A() As LxVF) As LxVFs
A_Ay = A
Set Init = Me
End Function
Function Ly() As String()
Dim O$(), J%
For J = 0 To UB(A_Ay)
    Push O, A_Ay(J).Lin
Next
End Function
Function Ay() As LxVF()
Ay = A_Ay
End Function
