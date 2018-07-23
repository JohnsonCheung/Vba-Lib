Attribute VB_Name = "M_DaoTy"
Option Explicit
Function DaoTy_Str$(T As DataTypeEnum)
Dim O$
Select Case T
Case Dao.DataTypeEnum.dbBoolean: O = "Boolean"
Case Dao.DataTypeEnum.dbDouble: O = "Double"
Case Dao.DataTypeEnum.dbText: O = "Text"
Case Dao.DataTypeEnum.dbDate: O = "Date"
Case Dao.DataTypeEnum.dbByte: O = "Byte"
Case Dao.DataTypeEnum.dbInteger: O = "Int"
Case Dao.DataTypeEnum.dbLong: O = "Long"
Case Dao.DataTypeEnum.dbDouble: O = "Doubld"
Case Dao.DataTypeEnum.dbDate: O = "Date"
Case Dao.DataTypeEnum.dbDecimal: O = "Decimal"
Case Dao.DataTypeEnum.dbCurrency: O = "Currency"
Case Dao.DataTypeEnum.dbSingle: O = "Single"
Case Else: Stop
End Select
DaoTy_Str = O
End Function
