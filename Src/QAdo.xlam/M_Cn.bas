Attribute VB_Name = "M_Cn"
Option Explicit
Function CnSqlDrs(A As Connection, Sql) As Drs
Set CnSqlDrs = RsDrs(A.Execute(Sql))
End Function

Sub CnRunSqlAy(A As Connection, SqlAy$())
If AyIsEmp(SqlAy) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   A.Execute Sql
Next
End Sub
