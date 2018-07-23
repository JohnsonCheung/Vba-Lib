Attribute VB_Name = "M_Dt"
Option Explicit
Sub DtInsDb(A As Database, Dt As Dt)
DbSqlAy_Run A, DbDt_SqlAy_OfIns(A, Dt)
End Sub
