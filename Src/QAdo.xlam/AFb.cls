VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AFb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$
Friend Property Get Init(Fb$) As AFb
A = Fb
Set Init = Me
End Property
Function ACnStr$()
ACnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
End Function

Sub Brw()
CurAcs.OpenCurrentDatabase A
CurAcs.Visible = True
End Sub
Property Get Cn() As Adodb.Connection
Dim O As New Adodb.Connection
O.Open ACnStr
Set Cn = O
End Property

Private Sub ZZ_SqlRun()
Const Sql$ = "Select * into [#a] from Permit"
If HasTbl("#a") Then SqlRun "Drop Table [#a]"
AFb(SampleFb_DutyPrepare).SqlRun Sql
End Sub
Function HasTbl(T) As Boolean
'HasTbl = Dbt(Db, T).Exist
Stop
End Function
Private Sub ZZ_SqlDrs()
Const Fb$ = SampleFb_DutyPrepare
Const Sql$ = "Select * from Permit"
AFb(Fb).SqlDrs(Sql).Brw
End Sub

Private Sub ZZ_Cn()
Dim A As Adodb.Connection
Set A = Init(SampleFb_DutyPrepare).Cn
Stop
End Sub

Property Get SqlARs(Sql) As ARs
Set SqlARs = ARs(SqlRs(Sql))
End Property

Property Get SqlRs(Sql) As Recordset
Set SqlRs = Cn.Execute(Sql)
End Property

Sub SqlRun(Sql$)
Cn.Execute Sql
End Sub

Property Get SqlDrs(Sql) As Drs
Set SqlDrs = ARs(SqlARs(Sql)).Drs
End Property

Friend Sub ZZ__Tst()
ZZ_Cn
ZZ_SqlRun
End Sub
