VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SyPair"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Sy1$(), B_Sy2$()
Property Get Sy1() As String()
Sy1 = B_Sy1
End Property
Property Get Sy2() As String()
Sy2 = B_Sy2
End Property
Property Let Sy1(V$())
B_Sy1 = V
End Property
Property Let Sy2(V$())
B_Sy2 = V
End Property
Property Get Init(Sy1$(), Sy2$()) As SyPair
B_Sy1 = Sy1
B_Sy2 = Sy2
Set Init = Me
End Property
Property Get N&()
N = Sz(B_Sy1)
End Property
Property Get U&()
U = N - 1
End Property
Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property
Property Get S1S2s() As S1S2s
If IsEmp Then Exit Function
Dim O As New S1S2s
Dim J&
For J = 0 To U
    O.Push S1S2(B_Sy1(J), B_Sy2(J))
Next
Set S1S2s = O
End Property
Property Get IsEqSz() As Boolean
IsEqSz = Sz(B_Sy1) = Sz(B_Sy2)
End Property

