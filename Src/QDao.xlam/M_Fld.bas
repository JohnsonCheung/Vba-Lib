Attribute VB_Name = "M_Fld"
Option Explicit
Function FldDes$(A As Dao.Field)
FldDes = PrpVal(A.Properties, "Description")
End Function
