Attribute VB_Name = "M_Cat"
Option Explicit
Property Get CatHasTbl(A As Catalog, T) As Boolean
CatHasTbl = ItrHasNm(A.Tables, T)
End Property
Property Get CatTny(A As Catalog, Optional Patn$ = ".") As String()
CatTny = ItrNy(A.Tables, Patn)
End Property

