Attribute VB_Name = "M_CnoValAy"
Option Explicit

Function CnoValAy_CnoAy(A() As CnoVal) As Integer()
CnoValAy_CnoAy = OyPrpIntAy(A, "Cno")
End Function

Function CnoValAy_CnoIx%(A() As CnoVal, Cno%)
'Use Cno to find any element in B_Ay has .Cno = Cno,
'Return the Ix of B_Ay if found else return -1
Dim J%
For J = 0 To UB(A)
    If A(J).Cno = Cno Then CnoValAy_CnoIx = J: Exit Function
Next
CnoValAy_CnoIx = -1
End Function

Function CnoValAy_StrValAy(A() As CnoVal) As String()
CnoValAy_StrValAy = OyPrpSy(A, "V")
End Function

Function CnoValAy_ToStr$(A() As CnoVal)
CnoValAy_ToStr = Tag("CnoValAy", OyToStr(A))
End Function

Function CnoValAy_ValAy(A() As CnoVal) As Variant()
CnoValAy_ValAy = OyPrpAy(A, "Val")
End Function
