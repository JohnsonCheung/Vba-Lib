Attribute VB_Name = "Module1"
Option Explicit

Sub AA()
Dim A() As ABC
ReDim A(0)
Set A(0) = New ABC
Dim B
B = A
Debug.Print VarPtr(B(0))
Debug.Print VarPtr(A(0))
Stop
End Sub
