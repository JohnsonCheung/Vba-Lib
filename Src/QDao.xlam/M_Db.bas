Attribute VB_Name = "M_Db"
Option Explicit

Property Get DbDs(A As Database, Tny0, Optional DsNm$ = "Ds") As Ds
Dim DtAy1() As Dt
    Dim U%, Tny$()
    Tny = DftNy(Tny0)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        DtAy(J) = DbtDt(A, Tny(J))
    Next
Set DbDs = Ds(DsNm, DtAy1)
End Property

Property Get DbInf(A As Database) As DbInf
Stop '
End Property

Property Get Ds(Tny0) As Ds
Dim Tny$(): Tny = DftNy(Tny)
End Property

Sub DbCrtTbl(A As Database, T, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Sub DbEnsTmp1Tbl(A As Database)
If DbHasTbl(A, "Tmp1") Then Exit Sub
DbqRun A, "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub

Function DbHasTbl(A As Database, T) As Boolean
DbHasTbl = A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Sub DbLnkFb(A As Database, Fb$, Tny0, Optional SrcTny0)
Dim Tny$(): Tny = DftNy(Tny0)              ' Src_Tny
Dim Src$(): Src = DftNy(Dft(SrcTny0, Tny0)) ' Tar_Tny
Ass Sz(Tny) > 0
Ass AyPair_IsEqSz(Src, Tny)
Dim J%
For J = 0 To UB(Tny)
    Dbt(A, Tny(J)).LnkFb Fb, Src(J)
Next
End Sub

Sub DbLnkFx(A As Database, Fx$, Optional WsNy0)
Dim WsNy$(): WsNy = Xls.Fx(Fx).DftWsNy(WsNy0)
Dim J%
For J = 0 To UB(WsNy)
   Dbt(A, WsNy(J)).LnkFxWs Fx
Next
End Sub

Function DbQny(A As Database) As String()
DbQny = DbqSy(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Private Sub ZZ_DbDs()
Dim Ds As Ds
Set Ds = DbDs(CurDb, "Permit PermitD")
Stop
End Sub

Private Sub ZZ_Qny()
AyDmp DbQny(CurDb)
End Sub
Property Get DbEng() As Dao.DBEngine
Static Y As New Dao.DBEngine
Set DbEng = Y
End Property
Sub DbRunSqlAy(A As Database, SqlAy$())
If AyIsEmp(A) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   DbqRun A, CStr(Sql)
Next
End Sub
Function DbTny(A As Database) As String()
DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'")
End Function
Function DbInfDs(A As Database) As Ds
Dim O As Ds
DsAddDt O, LnkDt
DsAddDt O, StruDt
DsAddDt O, TblFDt
DsAddDt O, PrpDt
O.DsNm = A.Name
Ds = O
End Function
Sub ZZ_DbInfBrw()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
DbInf(SampleDb_DutyPrepare).Brw
End Sub
Function DbLnkInfDt(A As Database) As Dt
Dim T, Dry(), C$
For Each T In DbTny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
LnkDt = NewDt("Lnk", DftNy("Tbl Connect"), Dry)
End Function
Function DbPrpInfDt(A As Database) As Dt
Set DbPrpInfDt = Dt("DbPrp", SplitSpc("A A"), Emp.Ay)
End Function
Function DbInfWb(A As Database, Optional Hid As Boolean) As Workbook
Dim O As Workbook
Set O = DsWb(Ds)
If Not Hid Then WbVis O
Set Wb = O
End Function
Function DbTblFInfDt(A As Database) As Dt
Dim T, Dry()
For Each T In DbTny(A)
   PushAy Dry, DbtTblFInfDry(A, T)
Next
Set DbTblFInfDt = Dt("TblFld", FnyOf_InfOf_TblF, Dry)
End Function
Sub DbBrwInf(A As Database)
AyBrw DsLy(DbInfDs(A), 2000, DtBrkLinMapStr:="TblFld:Tbl")
End Sub
Function DbStruInfDt(A As Database) As Dt
Dim T, Dry()
For Each T In DbTny(A)
    Push Dry, Array(T, DbtRecCnt(A, T), DbtDes(A, T), DbtStruLin(A, T, SkipTn:=True))
Next
Set DbStruInfDt = Dt("Tbl", "Tbl RecCnt Des Stru", Dry)
End Function
