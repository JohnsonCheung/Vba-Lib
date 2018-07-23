Attribute VB_Name = "M_Fny"
Option Explicit

Function FnyIxAy(A$(), SubFny0) As Integer()
Dim SubFny$(): SubFny = DftNy(SubFny0)
If AyIsEmp(SubFny) Then Stop
Dim O%(), U&, J%
U = UB(SubFny)
ReSz O, U
For J = 0 To U
    O(J) = AyIx(A, SubFny(J))
    If O(J) = -1 Then Stop
Next
End Function
