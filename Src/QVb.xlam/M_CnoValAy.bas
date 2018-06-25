Attribute VB_Name = "M_CnoValAy"
Option Explicit

Property Get CnoValAy_CnoAy(A() As CnoVal) As Integer()
CnoValAy_CnoAy = OyPrpIntAy(A, "Cno")
End Property

Property Get CnoValAy_CnoIx%(A() As CnoVal, Cno%)
'Use Cno to find any element in B_Ay has .Cno = Cno,
'Return the Ix of B_Ay if found else return -1
Dim J%
For J = 0 To UB(A)
    If A(J).Cno = Cno Then CnoValAy_CnoIx = J: Exit Property
Next
CnoValAy_CnoIx = -1
End Property

Property Get CnoValAy_StrValAy(A() As CnoVal) As String()
CnoValAy_StrValAy = OyPrpSy(A, "V")
End Property

Property Get CnoValAy_ToStr$(A() As CnoVal)
CnoValAy_ToStr = Tag("CnoValAy", OyToStr(A))
End Property

Property Get CnoValAy_ValAy(A() As CnoVal) As Variant()
CnoValAy_ValAy = OyPrpAy(A, "Val")
End Property
