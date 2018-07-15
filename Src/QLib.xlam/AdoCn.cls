VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoCn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Cn As ADODB.Connection

Sub RunSqlAy(SqlAy$())
If AyIsEmp(SqlAy) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   Cn.Execute CStr(Sql)
Next
End Sub

Function SqlDrs(Sql) As Drs
SqlDrs = Rs(Sql).Drs
End Function

Private Sub SqlDrs__Tst()
Dim Cn As ADODB.Connection: Set Cn = Fx(SampleFx_KE24).Cn
Dim Sql$: Sql = "Select * from [Sheet1$]"
Dim Drs As Drs: Drs = AdoP.Cn(Cn).SqlDrs(Sql)
DrsBrw Drs
End Sub

Private Function Rs(Sql) As AdoRs
Set Rs = AdoP.Rs(Cn.Execute(Sql))
End Function
