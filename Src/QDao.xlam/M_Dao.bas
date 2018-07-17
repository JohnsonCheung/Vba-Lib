Attribute VB_Name = "M_Dao"
Option Explicit
Public Const SampleFb_DutyPrepare$ = "C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"


Property Get HasFld(T As TableDef, F) As Boolean
'TblHasFld = FldsHasFld(T.Fields, F)
End Property

Property Get PrpVal(A As Dao.Properties, PrpNm$)
On Error Resume Next
PrpVal = A(PrpNm).Value
End Property

Property Get SampleDb_DutyPrepare() As Database
Set SampleDb_DutyPrepare = DbEng.OpenDatabase(SampleFb_DutyPrepare)
End Property

Property Get Tst() As DaoTst
Set Tst = New DaoTst
End Property

Function CurDb() As Database
Stop
End Function

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


Sub DbSqlAy_Run(A As Database, SqlAy$())
If AyIsEmp(A) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   DbqRun A, CStr(Sql)
Next
End Sub

Function DbTF_Fld(A As Database, T$, F) As Dao.Field
Set DbTF_Fld = A.TableDefs(T).Fields(F)
End Function

Function DbTF_FldInfDr(A As Database, T, F) As Variant()
Dim FF  As Dao.Field
Set FF = A.TableDefs(T).Fields(F)
With FF
    DbTF_FldInfDr = Array(F, IIf(DbTF_IsPk(A, T, F), "*", ""), DaoTy_Str(.Type), .Size, .DefaultValue, .Required, FldDes(FF))
End With
End Function

Function DbTF_IsPk(A As Database, T, F) As Boolean
DbTF_IsPk = AyHas(DbtPk(A, T), F)
End Function

Function DbTF_NxtId&(A As Database, T, Optional F)
Dim S$: S = FmtQQ("select Max(?) from ?", Dft(F, T), T)
DbTF_NxtId = DbqV(A, S) + 1
End Function

Function DbTny(A As Database) As String()
DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'")
End Function







Function DftDb(A As Database) As Database
If IsNothing(A) Then
   Set DftDb = CurDb
Else
   Set DftDb = A
End If
End Function

Function DftFb$(A$)
If A = "" Then
   Dim O$: O = TmpFb
   Dao.DBEngine.CreateDatabase(O, dbLangGeneral).Close
   DftFb = O
Else
   DftFb = A
End If
End Function

Function FnyOf_FldInf() As String()
FnyOf_FldInf = SplitSpc("Fld Pk Ty Sz Dft Req Des")
End Function

Function FnyOf_TblFInf() As String()
Dim O$()
Push O, "Tbl"
Push O, "SeqNo"
PushAy O, FnyOf_FldInf
FnyOf_TblFInf = O
End Function

Function FnyQuote(Fny$(), ToQuoteFny$()) As String()
If AyIsEmp(Fny) Then Exit Function
Dim O$(): O = Fny
Dim J%, F
For Each F In O
    If AyHas(ToQuoteFny, F) Then O(J) = Quote(CStr(F), "[]")
    J = J + 1
Next
FnyQuote = O
End Function

Function FnyQuoteIfNeed(Fny$()) As String()
If AyIsEmp(Fny) Then Exit Function
Dim O$(), J%, F
O = Fny
For Each F In Fny
    If IsNeedQuote(F) Then O(J) = Quote(CStr(F), "'")
    J = J + 1
Next
FnyQuoteIfNeed = O
End Function

Function FxTmpDb(Fx$, Optional WsNy0) As Database
Dim O As Database
   Set O = TmpDb
   DbLnkFx O, Fx, WsNy0
Set FxTmpDb = O
End Function

Function NewDb(Optional Fb0$, Optional Lang$ = Dao.LanguageConstants.dbLangGeneral) As Database
Dim Fb$
    Fb = DftFb(Fb0)
Ass Not FfnIsExist(Fb)
Set NewDb = Dao.DBEngine.CreateDatabase(Fb, Lang)
End Function

Function RsDry(A As Recordset) As Variant()
Dim O()
With A
   While Not .EOF
       Push O, FldsDr(A.Fields)
       .MoveNext
   Wend
End With
RsDry = O
End Function

Function TmpDb(Optional Fnn$) As Database
Set TmpDb = DBEngine.CreateDatabase(TmpFb("TmpDb", Fnn), Dao.LanguageConstants.dbLangGeneral)
End Function

Private Sub FxTmpDb__Tst()
Dim Db As Database: Set Db = FxTmpDb("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
AyDmp DbTny(Db)
Db.Close
End Sub

Private Sub ZZ_NewDb()
Dim A As Database
Set A = NewDb
Stop
End Sub
