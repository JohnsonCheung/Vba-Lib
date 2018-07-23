Attribute VB_Name = "M_Inf"
Option Explicit
Function InfDrs(Optional MdNm$, Optional Lno) As Drs
With InfDrs
    .Fny = InfFny
    .Dry = Array(InfDr(MdNm, Lno))
End With
End Function
