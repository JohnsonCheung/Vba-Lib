VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1S2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public S1$, S2$
Friend Property Get Init(A1, A2) As S1S2
S1 = A1
S2 = A2
Set Init = Me
End Property
Sub Asg(O1$, O2$)
O1 = S1
O2 = S2
End Sub
Property Get ToStr$()
ToStr = "S1(" & S1 & ") S2(" & S2 & ")"
End Property
Function Lin(Sep$, W1%)
Lin = AlignL(S1, W1) & Sep & S2
End Function

Function Ly(W1%, W2%) As String()
Dim Sy1$(), Sy2$()
Sy1 = SplitCrLf(S1)
Sy2 = SplitCrLf(S2)
Dim M%, J%, O$(), Lin$, A1$, A2$, U1%, U2%
    U1 = UB(Sy1)
    U2 = UB(Sy2)
    M = Max(U1, U2)
Dim Spc1$, Spc2$
    Spc1 = Space(W1)
    Spc2 = Space(W2)
For J = 0 To M
   If J > U1 Then
       A1 = Spc1
   Else
       A1 = AlignL(Sy1(J), W1)
   End If
   If J > U2 Then
       A2 = Spc2
   Else
       A2 = AlignL(Sy2(J), W2)
   End If
   Lin = "| " + A1 + " | " + A2 + " |"
   Push O, Lin
Next
Ly = O
End Function

