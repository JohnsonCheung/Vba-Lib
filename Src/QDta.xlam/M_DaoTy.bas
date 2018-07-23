Attribute VB_Name = "M_DaoTy"
Option Explicit

Function DaoTy_SimTy(A As DataTypeEnum) As eSimTy
Dim O As eSimTy
Select Case A
Case _
   DAO.DataTypeEnum.dbBigInt, _
   DAO.DataTypeEnum.dbByte, _
   DAO.DataTypeEnum.dbCurrency, _
   DAO.DataTypeEnum.dbDecimal, _
   DAO.DataTypeEnum.dbDouble, _
   DAO.DataTypeEnum.dbFloat, _
   DAO.DataTypeEnum.dbInteger, _
   DAO.DataTypeEnum.dbLong, _
   DAO.DataTypeEnum.dbNumeric, _
   DAO.DataTypeEnum.dbSingle
   O = eNbr
Case _
   DAO.DataTypeEnum.dbChar, _
   DAO.DataTypeEnum.dbGUID, _
   DAO.DataTypeEnum.dbMemo, _
   DAO.DataTypeEnum.dbText
   O = eTxt
Case _
   DAO.DataTypeEnum.dbBoolean
   O = eLgc
Case _
   DAO.DataTypeEnum.dbDate, _
   DAO.DataTypeEnum.dbTimeStamp, _
   DAO.DataTypeEnum.dbTime
   O = eDte
Case Else
   O = eOth
End Select
DaoTy_SimTy = O
End Function
