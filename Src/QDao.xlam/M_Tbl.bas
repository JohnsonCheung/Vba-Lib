Attribute VB_Name = "M_Tbl"
Option Explicit
Property Get TblHasFld(T As TableDef, F) As Boolean
TblHasFld = FldsHasFld(T.Fields, F)
End Property
