Attribute VB_Name = "M_S1S2"
Option Explicit

Property Get S1S2_Clone(A As S1S2) As S1S2
Set S1S2_Clone = S1S2(A.S1, A.S2)
End Property

Property Get S1S2_Lin$(A As S1S2, Optional Sep$ = " ", Optional W1%)
S1S2_Lin = AlignL(A.S1, W1) & Sep & A.S2
End Property
