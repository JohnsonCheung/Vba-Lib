Attribute VB_Name = "M_Db"
Option Explicit

Sub DbBrwInf(A As Database)
AyBrw DsLy(DbInfDs(A), 2000, DtBrkLinMapStr:="TblFld:Tbl")
End Sub

Sub DbCrtTbl(A As Database, T, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Function DbDs(A As Database, Tny0, Optional DsNm$) As Ds
Dim DtAy1() As Dt
    Dim U%, Tny$()
    Tny = DftNy(Tny0)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        Set DtAy(J) = DbtDt(A, Tny(J))
    Next
Set DbDs = Ds(DtAy1, DftDbNm(DsNm, A))
End Function

Sub DbEnsTmp1Tbl(A As Database)
If DbHasTbl(A, "Tmp1") Then Exit Sub
DbqRun A, "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub

Function DbHasTbl(A As Database, T) As Boolean
DbHasTbl = A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Function DbInfDs(A As Database) As Ds
Dim O As Ds
DsAddDt O, DbInfDtzLnk(A)
DsAddDt O, DbInfDtzStru(A)
DsAddDt O, DbInfDtzTblF(A)
DsAddDt O, DbInfDtzPrp(A)
O.DsNm = A.Name
DbInfDs = O
End Function

Function DbInfWb(A As Database) As Workbook
Set DbInfWb = DsWb(DbInfDs(A))
End Function

Sub DbLnkFb(A As Database, Fb$, Tny0, Optional SrcTny0)
Dim Tny$(): Tny = DftNy(Tny0)              ' Src_Tny
Dim Src$(): Src = DftNy(Dft(SrcTny0, Tny0)) ' Tar_Tny
Ass Sz(Tny) > 0
Ass Sz(Src) = Sz(Tny)
Dim J%
For J = 0 To UB(Tny)
    DbtLnkFb A, Tny(J), Fb, Src(J)
Next
End Sub

Sub DbLnkFx(A As Database, Fx$, Optional WsNy0)
Dim WsNy$(): WsNy = Xls.Fx(Fx).DftWsNy(WsNy0)
Dim J%
For J = 0 To UB(WsNy)
   DbtLnkFxWs A, Fx, WsNy(J)
Next
End Sub

Function DbInfDtzLnk(A As Database) As Dt
Dim T, Dry(), C$
For Each T In DbTny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
DbInfDtzLnk = Dt("Lnk", DftNy("Tbl Connect"), Dry)
End Function

Function DbInfDtzPrp(A As Database) As Dt
Set DbInfDtzPrp = Dt("DbPrp", SplitSpc("A A"), EmpAy)
End Function

Function DbQny(A As Database) As String()
DbQny = DbqSy(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Sub DbRunSqlAy(A As Database, SqlAy$())
If AyIsEmp(A) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   DbqRun A, CStr(Sql)
Next
End Sub

Function DbInfDtzStru(A As Database) As Dt
Dim T, Dry()
For Each T In DbTny(A)
    Push Dry, Array(T, DbtRecCnt(A, T), DbtDes(A, T), DbtStruLin(A, T, SkipTn:=True))
Next
Set DbInfDtzStru = Dt("Tbl", "Tbl RecCnt Des Stru", Dry)
End Function

Function DbInfDtzTblF(A As Database) As Dt
Dim T, Dry()
For Each T In DbTny(A)
   PushAy Dry, DbtInfDryzTblF(A, T)
Next
Set DbInfDtzTblF = Dt("TblFld", FnyzDbInfzTblF, Dry)
End Function

Function DbTny(A As Database) As String()
DbTny = DbqSy(A, "Select Name from MSysObjects where Type in (1,6) and Left(Name,4)<>'MSYS'")
End Function

Private Sub ZZ_DbDs()
Dim Ds As Ds
Dim Db As Database: Set Db = Sample_Db_DutyPrepare
Set Ds = DbDs(Db, "Permit PermitD")
DsBrw Ds
End Sub

Sub ZZ_DbInfBrw()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
DsBrw DbInfDs(Sample_Db_DutyPrepare)
End Sub

Private Sub ZZ_DbQny()
AyDmp DbQny(Sample_Db_DutyPrepare)
End Sub
