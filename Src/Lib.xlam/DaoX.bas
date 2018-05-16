Attribute VB_Name = "DaoX"
Option Explicit
Public Const SampleFb_DutyPrepare$ = "C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
Function SampleDb_DutyPrepare() As Database
Set SampleDb_DutyPrepare = FbDb(SampleFb_DutyPrepare)
End Function
Sub DbEnsTmp1Tbl(A As Database)
If DbHasTbl(A, "Tmp1") Then Exit Sub
DbqRun A, "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub
Function DbHasTbl(A As Database, T) As Boolean
DbHasTbl = A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Function RsDrs(A As Recordset) As Drs
RsDrs.Fny = RsFny(A)
RsDrs.Dry = RsDry(A)
End Function

Function FldsDr(A As dao.Fields) As Variant()
Dim O(), J%
ReDim O(A.Count - 1)
For J = 0 To A.Count - 1
   O(J) = A(J).Value
Next
FldsDr = O
End Function

Function FldsFny(A As dao.Fields) As String()
Dim O$(), J%
ReDim O(A.Count - 1)
For J = 0 To A.Count - 1
   O(J) = A(J).Name
Next
FldsFny = O
End Function

Function FldsHasFld(A As dao.Fields, F) As Boolean
Dim I  As dao.Field
For Each I In A
   If I.Name = F Then FldsHasFld = True: Exit Function
Next
End Function
Function FldDes$(A As dao.Field)
FldDes = PrpVal(A.Properties, "Description")
End Function

Sub DbLnkFb(A As Database, Fb$, Tny0, Optional SrcTny0)
Dim Tny$(): Tny = DftNy(Tny0)              ' Src_Tny
Dim Src$(): Src = DftNy(Dft(SrcTny0, Tny0)) ' Tar_Tny
Ass Sz(Tny) > 0
Ass AyPair_IsEqSz(Src, Tny)
Dim J%
For J = 0 To UB(Tny)
    DbtLnkFb A, Tny(J), Fb, Src(J)
Next
End Sub
Sub DbtLnkFb(A As Database, T, Fb$, Optional SrcT0$)
Dim Src$: Src = Dft(SrcT0, T)
Dim Tbl  As TableDef
Set Tbl = A.CreateTableDef(T)
Tbl.SourceTableName = Src
Tbl.Connect = ";DATABASE=?" & Fb
DbtDrp A, T
A.TableDefs.Append Tbl
End Sub
Sub DbLnkFx(A As Database, Fx$, Optional WsNy0)
Dim WsNy$(): WsNy = DftNy(WsNy0)
WsNy = DftFxWsNy(Fx, WsNy)
Dim J%
For J = 0 To UB(WsNy)
   DbtLnkFxWs A, WsNy(J), Fx
Next
End Sub

Function DbtFxOfLnkTbl$(A As Database, T)
DbtFxOfLnkTbl = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Sub DbtLnkFxWs(A As Database, T$, Fx$, Optional WsNm0)
Const CSub$ = "ATLnkFxWs"
Dim WsNm$: WsNm = Dft(WsNm0, T)
On Error GoTo X
   Dim Tbl  As TableDef
   Set Tbl = A.CreateTableDef(T)
   Tbl.SourceTableName = WsNm & "$"
   Tbl.Connect = FmtQQ("Excel 8.0;HDR=YES;IMEX=2;DATABASE=?", Fx)
   DbtDrp A, T
   A.TableDefs.Append Tbl
Exit Sub
X: Er CSub, "{Er} found in Creating {T} in {Db} by Linking {WsNm} in {Fx}", Err.Description, T, A.Name, WsNm0, Fx
End Sub

Function DftFxWsNy(Fx$, WsNy$()) As String()
If AyIsEmp(WsNy) Then
   DftFxWsNy = FxWsNy(Fx)
   Exit Function
End If
DftFxWsNy = WsNy
End Function

Function DftNy(Ny0) As String()
If ValIsStr(Ny0) Then
   DftNy = LvsSy(Ny0)
   Exit Function
End If
If ValIsSy(Ny0) Then
   DftNy = Ny0
End If
End Function

Sub DbSqlAy_Run(A As Database, SqlAy$())
If AyIsEmp(A) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   DbqRun A, CStr(Sql)
Next
End Sub

Function DbqDrs(A As Database, Sql$) As Drs
Dim Rs As Recordset
Dim O As Drs
Set Rs = A.OpenRecordset(Sql)
O.Dry = RsDry(Rs)
O.Fny = RsFny(Rs)
DbqDrs = O
End Function

Sub DbtAddFld(A As Database, T, F, Ty As DataTypeEnum)
Dim FF As New dao.Field
FF.Name = F
FF.Type = Ty
DbtFlds(A, T).Append FF
End Sub

Function DbtDt(A As Database, T) As Dt
Dim O As Dt
O.DtNm = T
O.Dry = RsDry(A.TableDefs(T).OpenRecordset)
O.Fny = DbtFny(A, T)
DbtDt = O
End Function

Function DbtExist(A As Database, T) As Boolean
DbtExist = A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Function DbtFlds(A As Database, T) As dao.Fields
Set DbtFlds = A.TableDefs(T).Fields
End Function

Function DbtFny(A As Database, T) As String()
DbtFny = FldsFny(A.TableDefs(T).Fields)
End Function

Function DbtPk(A As Database, T) As String()
Dim I  As Index, O$(), F
On Error GoTo X
If A.TableDefs(T).Indexes.Count = 0 Then Exit Function
On Error GoTo 0
For Each I In A.TableDefs(T).Indexes
   If I.Primary Then
       For Each F In I.Fields
           Push O, F.Name
       Next
       DbtPk = O
       Exit Function
   End If
Next
X:
End Function

Function DbtRecCnt&(A As Database, T)
DbtRecCnt = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
End Function

Function DbtSimTyAy(A As Database, T$, Optional Fny0) As eSimTy()
Dim Fny$(): Fny = DftFny(A, T, Fny0)
Dim O() As eSimTy
   Dim U%
   ReDim O(U)
   Dim J%, F
   J = 0
   For Each F In Fny
       O(J) = NewSimTy(DbTF_Fld(A, T, F).Type)
       J = J + 1
   Next
DbtSimTyAy = O
End Function

Function DbtWs(A As Database, T) As Worksheet
Set DbtWs = DtWs(DbtDt(A, T))
End Function

Function DbqDry(A As Database, Sql$) As Variant()
DbqDry = RsDry(A.OpenRecordset(Sql))
End Function

Sub DbqRun(A As Database, Sql$)
A.Execute Sql
End Sub

Function DbqSy(A As Database, Sql$) As String()
DbqSy = RsSy(A.OpenRecordset(Sql))
End Function

Function DbqV(A As Database, Sql$)
With A.OpenRecordset(Sql)
   DbqV = .Fields(0).Value
   .Close
End With
End Function

Function DtaDb() As Database
Set DtaDb = DBEngine.OpenDatabase(DtaFb)
End Function

Function DtaFb$()
DtaFb = FfnRplExt(FfnAddFnSfx(CurFb, "_Data"), ".mdb")
End Function

Function FbAcnStr$(A$)
FbAcnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
End Function

Sub FbBrw(Fb$)
CurAcs.OpenCurrentDatabase Fb
CurAcs.Visible = True
End Sub
Function FbDb(A$) As Database
Set FbDb = dao.DBEngine.OpenDatabase(A)
End Function
Function FbCn(A$) As ADODB.Connection
Dim O As New ADODB.Connection
O.Open FbAcnStr(A)
Set FbCn = O
End Function

Private Sub FbSql_Arun__Tst()
Const Fb$ = "N:\SapAccessReports\DutyPrepay5\tmp.accdb"
Const Sql$ = "Select * into [#a] from Permit"
FbSql_Arun Fb, "Drop Table [#a]"
FbSql_Arun Fb, Sql
End Sub

Function FxTmpDb(Fx$, Optional WsNy0) As Database
Dim O As Database
   Set O = TmpDb
   DbLnkFx O, Fx, WsNy0
Set FxTmpDb = O
End Function

Private Sub FxTmpDb__Tst()
Dim Db As Database: Set Db = FxTmpDb("N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls")
AyDmp DbTny(Db)
Db.Close
End Sub

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

Function RsFny(A As Recordset) As String()
RsFny = FldsFny(A.Fields)
End Function

Function RsSy(A As Recordset) As String()
Dim O$()
With A
   While Not .EOF
       Push O$, A.Fields(0).Value
       .MoveNext
   Wend
End With
RsSy = O
End Function

Function DftFb$(A$)
If A = "" Then
   Dim O$: O = TmpFb
   dao.DBEngine.CreateDatabase(O, dbLangGeneral).Close
   DftFb = O
Else
   DftFb = A
End If
End Function

Function DftFny(A As Database, T$, Fny) As String()
If IsMissing(Fny) Then
   DftFny = DbtFny(A, T)
Else
   DftFny = Fny
End If
End Function

Function DftDb(A As Database) As Database
If IsNothing(A) Then
   Set DftDb = CurDb
Else
   Set DftDb = A
End If
End Function

Function NewDb(Optional Fb0$, Optional Lang$ = dao.LanguageConstants.dbLangGeneral) As Database
Dim Fb$
    Fb = DftFb(Fb0)
Ass Not FfnIsExist(Fb)
Set NewDb = dao.DBEngine.CreateDatabase(Fb, Lang)
End Function

Private Sub NewDb__Tst()
Dim A As Database
Set A = NewDb
Stop
End Sub

Function DaoTy_Str$(T As DataTypeEnum)
Dim O$
Select Case T
Case dao.DataTypeEnum.dbBoolean: O = "Boolean"
Case dao.DataTypeEnum.dbDouble: O = "Double"
Case dao.DataTypeEnum.dbText: O = "Text"
Case dao.DataTypeEnum.dbDate: O = "Date"
Case dao.DataTypeEnum.dbByte: O = "Byte"
Case dao.DataTypeEnum.dbInteger: O = "Int"
Case dao.DataTypeEnum.dbLong: O = "Long"
Case dao.DataTypeEnum.dbDouble: O = "Doubld"
Case dao.DataTypeEnum.dbDate: O = "Date"
Case dao.DataTypeEnum.dbDecimal: O = "Decimal"
Case dao.DataTypeEnum.dbCurrency: O = "Currency"
Case dao.DataTypeEnum.dbSingle: O = "Single"
Case Else: Stop
End Select
DaoTy_Str = O
End Function

Sub DbBrw(A As Database)
Dim N$: N = A.Name
A.Close
FbBrw N
End Sub

Sub DbCrtTbl(A As Database, T, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Function DbQny(A As Database) As String()
DbQny = DbqSy(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Function DbTF_Fld(A As Database, T$, F) As dao.Field
Set DbTF_Fld = A.TableDefs(T).Fields(F)
End Function

Function DbTF_FldInfDr(A As Database, T, F) As Variant()
Dim FF  As dao.Field
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


Function DbT_HasFld(A As Database, T, F) As Boolean
Ass DbtExist(A, T)
DbT_HasFld = TblHasFld(A.TableDefs(T), F)
End Function

Function DbTny(A As Database) As String()
DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'")
End Function

Sub DbqBrw(A As Database, Sql$)
DrsBrw DbqDrs(A, Sql)
End Sub

Sub DbtBrw(A As Database, T)
DtBrw DbtDt(A, T)
End Sub

Function DbtDes$(A As Database, T)
DbtDes = PrpVal(A.TableDefs(T).Properties, "Description")
End Function

Sub DbtDrp(A As Database, T)
If DbHasTbl(A, T) Then A.Execute FmtQQ("Drop Table [?]", T)
End Sub

Function DbtTblFInfDry(A As Database, T) As Variant()
Dim O(), F, Dr(), Fny$()
Fny = DbtFny(A, T)
If AyIsEmp(Fny) Then Exit Function
Dim SeqNo%
SeqNo = 0
For Each F In Fny
    Erase Dr
    Push Dr, T
    Push Dr, SeqNo: SeqNo = SeqNo + 1
    PushAy Dr, DbTF_FldInfDr(A, T, CStr(F))
    Push O, Dr
Next
DbtTblFInfDry = O
End Function

Function FnyOf_InfOf_Fld() As String()
FnyOf_InfOf_Fld = SplitSpc("Fld Pk Ty Sz Dft Req Des")
End Function

Function FnyOf_InfOf_TblF() As String()
Dim O$()
Push O, "Tbl"
Push O, "SeqNo"
PushAy O, FnyOf_InfOf_Fld
FnyOf_InfOf_TblF = O
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

Function PrpVal(A As dao.Properties, PrpNm$)
On Error Resume Next
PrpVal = A(PrpNm).Value
End Function

Function TblHasFld(T As TableDef, F) As Boolean
'TblHasFld = FldsHasFld(T.Fields, F)
End Function

Function TmpDb(Optional Fnn$) As Database
Set TmpDb = DBEngine.CreateDatabase(TmpFb("TmpDb", Fnn), dao.LanguageConstants.dbLangGeneral)
End Function

Private Sub DbQny__Tst()
AyDmp DbQny(CurDb)
End Sub

Private Sub DbtPk__Tst()
Dim Dr(), Dry(), T
Dim Db As Database
Set Db = CurDb
For Each T In DbTny(Db)
    Erase Dr
    Push Dr, T
    PushAy Dr, DbtPk(Db, CStr(T))
    Push Dry, Dr
Next
DryBrw Dry
End Sub

Sub Tst__DaoDb()
DbtPk__Tst
End Sub
