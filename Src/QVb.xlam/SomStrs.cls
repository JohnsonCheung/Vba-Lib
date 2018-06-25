VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SomStrs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A() As SomStr

Property Get HasNone() As Boolean
Dim J%
For J = 0 To U
    If Not A(J).Som Then HasNone = True: Exit Property
Next
End Property

Property Get Init(Ay() As SomStr) As SomStrs
A = Ay
Set Init = Me
End Property

Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property

Property Get N&()
N = U - 1
End Property

Property Get U&()
U = N - 1
End Property
