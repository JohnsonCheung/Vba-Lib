Attribute VB_Name = "M_Cat"
Option Explicit
Function CatHasTbl(A As Catalog, T) As Boolean
CatHasTbl = ItrHasNm(A.Tables, T)
End Function
Function CatTny(A As Catalog, Optional Patn$ = ".") As String()
CatTny = ItrNy(A.Tables, Patn)
End Function

