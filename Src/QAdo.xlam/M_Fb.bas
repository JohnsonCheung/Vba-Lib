Attribute VB_Name = "M_Fb"
Option Explicit

Function FbCn(A) As Connection
Dim O As New Adodb.Connection
O.Open FbCnStr(A)
Set FbCn = O
End Function

Function FbCnStr$(A)
FbCnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;", A)
End Function

Function FbSqlRs(A, Sql) As Recordset
Set FbSqlRs = FbCn(A).Execute(Sql)
End Function

Function FbSqlDrs(A, Sql) As Drs
Set FbSqlDrs = RsDrs(FbCn(A).Execute(Sql))
End Function

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
AyDmp FbTny(Sample_Fb_DutyPrepare)
End Sub

Sub ZZ_FbHasTbl()
Ass FbHasTbl(Sample_Fb_DutyPrepare, "SkuB")
End Sub

Sub FbSqlRun(A, Sql)
FbCn(A).Execute Sql
End Sub
