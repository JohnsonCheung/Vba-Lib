Attribute VB_Name = "M_S1S2Ay"
Option Explicit

Function S1S2Ay_Drs(A() As S1S2) As Drs
Set S1S2Ay_Drs = Drs("S1 S2", S1S2Ay_Dry(A))
End Function

Function S1S2Ay_Dry(A() As S1S2) As Variant()
Dim O()
Dim J%
For J = 0 To UB(A)
   With A(J)
       Push O, Array(.S1, .S2)
   End With
Next
S1S2Ay_Dry = O
End Function
