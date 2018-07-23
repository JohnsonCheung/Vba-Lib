Attribute VB_Name = "M_Prp"
Option Explicit
Property Get PrpVal(A As Properties, PrpNm$)
On Error Resume Next
PrpVal = A(PrpNm).Value
End Property
