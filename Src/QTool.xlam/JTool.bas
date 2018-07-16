Attribute VB_Name = "JTool"
Option Explicit
Property Get Mth(A As CodeModule, MthNm) As Mth
Dim O As New Mth
Set Mth = O.Init(A, MthNm)
End Property
Property Get MthOpt(A As Mth) As MthOpt
Dim O As New MthOpt
Set MthOpt = O.Init(A)
End Property
Property Get S1S2(S1, S2) As S1S2
Dim O As New S1S2
Set S1S2 = O.Init(S1, S2)
End Property
