Attribute VB_Name = "M_Flds"
Option Explicit

Property Get FldsDr(A As Fields) As Variant()
FldsDr = ItrPrpValAy(A, "Value")
End Property

Property Get FldsFny(A As Fields) As String()
FldsFny = ItrNy(A)
End Property
