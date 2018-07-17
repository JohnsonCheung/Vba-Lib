Attribute VB_Name = "M_Dbq"
Option Explicit
Sub DbqBrw(A As Database, Sql$)
DrsBrw DbqDrs(A, Sql)
End Sub
Function DbqDrs(A As Database, Sql$) As Drs
Dim Rs As Recordset
Set Rs = A.OpenRecordset(Sql)
Set DbqDrs = Drs(RsFny(Rs), RsDry(Rs))
End Function
Function DbqDry(A As Database, Sql$) As Variant()
DbqDry = RsDry(A.OpenRecordset(Sql))
End Function
Sub DbqRun(A As Database, Sql$)
A.Execute Sql
End Sub
Function DbqSy(A As Database, Sql$) As String()
DbqSy = RsSy(A.OpenRecordset(Sql))
End Function
Function DbqV(A As Database, Sql$)
With A.OpenRecordset(Sql)
   DbqV = .Fields(0).Value
   .Close
End With
End Function
