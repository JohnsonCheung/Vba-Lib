VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Private A$
'Friend Property Get Init(Fb$) As Fb
'A = Fb
'Set Init = Me
'End Property
'Function ACnStr$()
'ACnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
'End Function
'
'Sub Brw()
'CurAcs.OpenCurrentDatabase A
'CurAcs.Visible = True
'End Sub
'Property Get Cn() As AdoDb.Connection
'Dim O As New AdoDb.Connection
'O.Open ACnStr
'Set Cn = O
'End Property
'
'Property Get Db() As Database
'Set Db = Dao.DBEngine.OpenDatabase(A)
'End Property
'
'Private Sub SqlARun__Tst()
'Const Sql$ = "Select * into [#a] from Permit"
'If HasTbl("#a") Then SqlArun "Drop Table [#a]"
'Fb(SampleFb_DutyPrepare).SqlArun Sql
'End Sub
'Function HasTbl(T) As Boolean
'HasTbl = Dbt(Db, T).Exist
'End Function
'Private Sub SqlADrs__Tst()
'Const Fb$ = SampleFb_DutyPrepare
'Const Sql$ = "Select * from Permit"
'DrsBrw Ado.Fb(Fb).SqlADrs(Sql)
'End Sub
'
'Private Sub Cn__Tst()
'Dim A As AdoDb.Connection
'Set A = Init(SampleFb_DutyPrepare).Cn
'Stop
'End Sub
'
'Function SqlARs(Sql) As AdoDb.Recordset
'Set SqlARs = Cn.Execute(Sql)
'End Function
'
'Sub SqlArun(Sql$)
'Cn.Execute Sql
'End Sub
'
'Function SqlADrs(Sql) As Drs
'SqlADrs = AdoP.Rs(SqlARs(Sql)).Drs
'End Function
'
'Friend Sub Tst()
'Cn__Tst
'SqlARun__Tst
'SqlADrs__Tst
'End Sub
'
