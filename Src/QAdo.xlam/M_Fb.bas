Attribute VB_Name = "M_Fb"
Option Explicit

Property Get FbCn(A) As Connection
Dim O As New Adodb.Connection
O.Open FbCnStr(A)
Set FbCn = O
End Property

Property Get FbCnStr$(A)
FbCnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
End Property

Property Get FbSqlRs(A, Sql) As Recordset
Set FbSqlRs = FbCn(A).Execute(Sql)
End Property

Property Get FbSqlDrs(A, Sql) As Drs
Set FbSqlDrs = RsDrs(FbCn(A).Execute(Sql))
End Property

Function FbCat(A) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = FbCn(A)
Set FbCat = O
End Function

Function FbHasTbl(A, T) As Boolean
FbHasTbl = CatHasTbl(FbCat(A), T)
End Function

Function FbTny(A, Optional Patn$ = ".") As String()
FbTny = CatTny(FbCat(A), Patn)
End Function

Sub ZZ_FbTny()
AyDmp FbTny(SampleFb_DutyPrepare)
End Sub

Sub ZZ_FbHasTbl()
Ass FbHasTbl(SampleFb_DutyPrepare, "SkuB")
End Sub

Sub FbSqlRun(A, Sql)
FbCn(A).Execute Sql
End Sub
