Attribute VB_Name = "M_Prj"
Option Explicit
Private Sub ZZ_PrjSrcDrs()
Dim O As Drs: O = CurPjx.SrcDrs
'DryBrw O

Dim A As SrcLin: Set A = V(O.Dry(2)(1)).SrcLin
Dim A1 As Drs: A1 = A.InfDrs
DrsDmp A1
Stop
End Sub
