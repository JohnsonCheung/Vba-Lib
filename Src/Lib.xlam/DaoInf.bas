Attribute VB_Name = "DaoInf"
Option Explicit

Sub DbBrwInf(A As Database)
AyBrw DsLy(DbInfDs(A), 2000, DtBrkLinMapStr:="TblFld:Tbl")
Exit Sub
WbVis DsWb(DbInfDs(A))
End Sub

Function DbInfDs(A As Database) As Ds
Dim O As Ds
DsAddDt O, DbInfDtOfLnk(A)
DsAddDt O, DbInfDtOfStru(A)
DsAddDt O, DbInfDtOfTblF(A)
DsAddDt O, DbInfDtOfPrp(A)
O.DsNm = A.Name
DbInfDs = O
End Function

Function DbInfDtOfLnk(A As Database) As Dt
Dim T, Dry(), C$
For Each T In DbTny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
DbInfDtOfLnk = NewDt("Lnk", DftNy("Tbl Connect"), Dry)
End Function

Function DbInfDtOfPrp(Optional D As Database) As Dt
DbInfDtOfPrp = NewDt("DbPrp", SplitSpc("A A"), EmpAy)
End Function

Function DbInfDtOfStru(A As Database) As Dt
Dim T, Dry(), TT$
For Each T In DbTny(A)
   TT = T
   Push Dry, Array(T, DbtRecCnt(A, TT), DbtDes(A, TT), DbtStruLin(A, TT, SkipTn:=True))
Next
Dim O As Dt
   With O
       .Dry = Dry
       .Fny = ApSy("Tbl", "RecCnt", "Des", "Stru")
       .DtNm = "Tbl"
   End With
DbInfDtOfStru = O
End Function

Function DbInfDtOfTblF(A As Database) As Dt
Dim T, Dry()
For Each T In DbTny(A)
   PushAy Dry, DbtTblFInfDry(A, T)
Next
Dim O As Dt
O.Dry = Dry
O.Fny = FnyOf_InfOf_TblF
O.DtNm = "TblFld"
DbInfDtOfTblF = O
End Function

Function DbtStruLin$(A As Database, T$, Optional SkipTn As Boolean)
Dim O$(): O = DbtFny(A, T): If AyIsEmp(O) Then Exit Function
O = FnyQuote(O, DbtPk(A, T))
O = FnyQuoteIfNeed(O)
Dim J%, V
V = 0
For Each V In O
   O(J) = Replace(V, T, "*")
   J = J + 1
Next
If SkipTn Then
   DbtStruLin = JnSpc(O)
Else
   DbtStruLin = T & " = " & JnSpc(O)
End If
End Function

Sub DbBrwInf__Tst()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
DbBrwInf FbDb(SampleFb_DutyPrepare)
End Sub
