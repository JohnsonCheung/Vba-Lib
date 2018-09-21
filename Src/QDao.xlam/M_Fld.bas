Attribute VB_Name = "M_Fld"
Option Explicit
Function FldDes$(A As DAO.Field)
FldDes = PrpVal(A.Properties, "Description")
End Function
