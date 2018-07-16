VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ACn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Cn As Adodb.Connection
Friend Property Get Init(A As Adodb.Connection) As ACn
Set Cn = A
Set Init = Me
End Property
Private Property Get Rs(Sql) As Recordset
Set Rs = Cn.Execute(Sql)
End Property
Private Property Get ARs(Sql) As ARs
Set ARs = ARs(Rs(Sql))
End Property
Function SqlDrs(Sql) As Drs
SqlDrs = ARs(Sql).Drs
End Function

Sub RunSqlAy(SqlAy$())
If AyIsEmp(SqlAy) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   Cn.Execute CStr(Sql)
Next
End Sub

Private Sub ZZ_SqlDrs()
Dim Sql$: Sql = "Select * from [Sheet1$]"
AFx(SampleFx_KE24).ACn.SqlDrs(Sql).Brw
End Sub
