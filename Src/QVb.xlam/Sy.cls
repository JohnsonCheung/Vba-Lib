VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$()

Property Get Init(Sy$()) As Sy
A = Sy
Set Init = Me
End Property

Function IsEmp() As Boolean
IsEmp = N = 0
End Function

Function IxLy() As String()
'It is 2 line first line is 0 ...
'first line is x0 x1 of A$()
If IsEmp Then Exit Function
Dim A1$()
Dim A2$()
ReSz A1, U
ReSz A2, U
Dim O$(), J%, L$, W%
For J = 0 To U
    L = Len(A(J))
    W = Max(L, Len(J))
    A1(J) = AlignL(J, W)
    A2(J) = AlignL(A(J), W)
Next
Push O, JnSpc(A1)
Push O, JnSpc(A2)
IxLy = O
End Function

Function N&()
On Error Resume Next
N = UBound(A) + 1
End Function

Function U&()
U = N - 1
End Function
