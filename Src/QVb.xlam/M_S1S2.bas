Attribute VB_Name = "M_S1S2"
Option Explicit

Function S1S2_Clone(A As S1S2) As S1S2
Set S1S2_Clone = S1S2(A.S1, A.S2)
End Function

Function S1S2_Lin$(A, Optional Sep$ = " ", Optional W1%)
S1S2_Lin = AlignL(A.S1, W1) & Sep & A.S2
End Function

Sub S1S2_Asg(A As S1S2, O1$, O2$)
O1 = A.S1
O2 = A.S2
End Sub
