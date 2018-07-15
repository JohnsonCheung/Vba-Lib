Attribute VB_Name = "M_DbInf"
Option Explicit

Sub DbBrwInf(A As Database)
AyBrw DsLy(DbInfDs(A), 2000, DtBrkLinMapStr:="TblFld:Tbl")
End Sub

Private Function TblFDt() As Dt
Dim T, Dry()
For Each T In DbTny(A)
   PushAy Dry, Me.Dbt(T).TblFInfDry
Next
Dim O As Dt
O.Dry = Dry
O.Fny = FnyOf_InfOf_TblF
O.DtNm = "TblFld"
TblFDt = O
End Function

Function DbInfWb(A As Database, Optional Hid As Boolean) As Workbook
Dim O As Workbook
Set O = DsWb(Ds)
If Not Hid Then WbVis O
Set Wb = O
End Function

Sub Brw__Tst()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
DbInf(SampleDb_DutyPrepare).Brw
End Sub

Function DbInfDs(A As Database) As Ds
Dim O As Ds
DsAddDt O, LnkDt
DsAddDt O, StruDt
DsAddDt O, TblFDt
DsAddDt O, PrpDt
O.DsNm = A.Name
Ds = O
End Function

Private Function LnkDt() As Dt
Dim T, Dry(), C$
For Each T In DbTny(A)
   C = A.TableDefs(T).Connect
   If C <> "" Then Push Dry, Array(T, C)
Next
Dim O As Dt
LnkDt = NewDt("Lnk", DftNy("Tbl Connect"), Dry)
End Function

Private Function PrpDt() As Dt
PrpDt = NewDt("DbPrp", SplitSpc("A A"), Emp.Ay)
End Function

Private Function StruDt() As Dt
Dim T, Dry(), TT$
For Each T In DbTny(A)
   TT = T
   With Me.Dbt(TT)
       Push Dry, Array(T, .RecCnt, .Des, .StruLin(SkipTn:=True))
    End With
Next
Dim O As Dt
   With O
       .Dry = Dry
       .Fny = ApSy("Tbl", "RecCnt", "Des", "Stru")
       .DtNm = "Tbl"
   End With
StruDt = O
End Function
