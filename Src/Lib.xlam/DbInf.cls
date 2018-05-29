VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DbInf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A As Dao.Database
Friend Function Init(Db As Database) As DbInf
Set A = Db
Set Init = Me
End Function

Friend Sub Tst()
Brw__Tst
End Sub
Sub Brw()
AyBrw DsLy(Ds, 2000, DtBrkLinMapStr:="TblFld:Tbl")
End Sub
Function Wb(Optional Hid As Boolean) As Workbook
Dim O As Workbook
Set O = DsWb(Ds)
If Not Hid Then WbVis O
Set Wb = O
End Function

Private Function Ds() As Ds
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
PrpDt = NewDt("DbPrp", SplitSpc("A A"), EmpAy)
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

Friend Function TblFDt() As Dt
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

Friend Function Dbt(T) As Dbt
Set Dbt = DaoX.Dbt(A, T)
End Function

Sub Brw__Tst()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlDim A As DBEngine: Set A = dao.DBEngine
'not work: dao.DBEngine.Workspaces(1).Databases(1).Execute "GRANT SELECT ON MSysObjects TO Admin;"
Me.Init SampleDb_DutyPrepare
Brw
End Sub


