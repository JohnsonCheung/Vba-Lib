Attribute VB_Name = "M_Flds"
Option Explicit

Function FldsDr(A As Fields) As Variant()
FldsDr = ItrPrpValAy(A, "Value")
End Function

Function FldsFny(A As Fields) As String()
FldsFny = ItrNy(A)
End Function
