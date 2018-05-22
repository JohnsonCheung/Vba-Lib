Attribute VB_Name = "Ado"
Option Explicit
Function AfldsDr(A As ADODB.Fields) As Variant()
Dim I As ADODB.Field
Dim O()
For Each I In A
   Push O, I.Value
Next
AfldsDr = O
End Function

Function AfldsFny(A As ADODB.Fields) As String()
Dim O$()
Dim F As ADODB.Field
For Each F In A
    Push O, F.Name
Next
AfldsFny = O
End Function

Function ArsDrs(A As ADODB.Recordset) As Drs
ArsDrs.Fny = ArsFny(A)
ArsDrs.Dry = ArsDry(A)
End Function

Function ArsDry(A As ADODB.Recordset) As Variant()
Dim O()
With A
    While Not .EOF
        Push O, AfldsDr(A.Fields)
        .MoveNext
    Wend
End With
ArsDry = O
End Function

Function ArsFny(A As ADODB.Recordset) As String()
ArsFny = AfldsFny(A.Fields)
End Function

Sub CnSqlAy_Run(A As ADODB.Connection, SqlAy$())
If AyIsEmp(SqlAy) Then Exit Sub
Dim Sql
For Each Sql In SqlAy
   A.Execute CStr(Sql)
Next
End Sub

Function CnSql_Drs(A As ADODB.Connection, Sql) As Drs
CnSql_Drs = ArsDrs(A.Execute(Sql))
End Function

Function DftWsNmByFxFstWs$(WsNm0, Fx)
Dim O$
If WsNm0 = "" Then O = FxFstWsNm(Fx) Else O = WsNm0
DftWsNmByFxFstWs = O
End Function

Function FbSql_Ars(A$, Sql) As ADODB.Recordset
Set FbSql_Ars = FbCn(A).Execute(Sql)
End Function

Sub FbSql_Arun(A$, Sql$)
FbCn(A).Execute Sql
End Sub

Function FbSql_Drs(A$, Sql$) As Drs
FbSql_Drs = ArsDrs(FbSql_Ars(A, Sql))
End Function

Function FxCat(A) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = FxCn(A)
Set FxCat = O
End Function

Function FxFstWsNm$(A)
Dim T As ADOX.Table
Dim O$()
For Each T In FxCat(A).Tables
    FxFstWsNm = RmvLasChr(T.Name)
    Exit Function
Next
End Function

Sub FxSql_Arun(A$, Sql)
FxCn(A).Execute CStr(Sql)
End Sub

Function FxSql_Drs(A$, Sql) As Drs
FxSql_Drs = ArsDrs(FxCn(A).Execute(Sql))
End Function

Function FxWsNy(A) As String()
FxWsNy = SyRmvLasChr(ItrNy(FxCat(A).Tables))
End Function

Function FxWs_Dt(A$, Optional WsNm0$) As Dt
Dim WsNm$
If WsNm0 = "" Then WsNm = FxFstWsNm(A) Else WsNm = WsNm0
Dim Sql$: Sql = FmtQQ("Select * from [?$]", WsNm)
FxWs_Dt = DtNmDrs_Dt(WsNm, CnSql_Drs(FxCn(A), Sql))
End Function

Function FxWs_Fny(A, Optional WsNm0$) As String()
Dim WsNm$: WsNm = DftWsNmByFxFstWs(WsNm, A)
FxWs_Fny = ItrNy(FxCat(A).Tables(WsNm & "$").Columns)
End Function

Sub ArsDry__Tst()
'strDdl = "GRANT SELECT ON MSysObjects TO Admin;"
'CurrentProject.Connection.Execute strDdlEnd Sub
DryBrw ArsDry(FbCn(SampleFb_DutyPrepare).Execute("Select * from KE24"))
End Sub

Private Sub CnSql_Drs__Tst()
Dim Cn As ADODB.Connection: Set Cn = FxCn(SampleFx_KE24)
Dim Sql$: Sql = "Select * from [Sheet1$]"
Dim Drs As Drs: Drs = CnSql_Drs(Cn, Sql)
DrsBrw Drs
End Sub

Private Sub FbAqlDrs__Tst()
Const Fb$ = "N:\SapAccessReports\DutyPrepay5\DutyPrepay5.accdb"
Const Sql$ = "Select * from Permit"
'DrsBrw FbAqlDrs(Fb, Sql)
End Sub

Private Sub FbCn__Tst()
Dim A As ADODB.Connection
Set A = FbCn("N:\SapAccessReports\DutyPrepay5\DutyPrepay5_data.mdb")
Stop
End Sub

Private Sub FxSql_Arun__Tst()
Const Fx$ = SampleFx_KE24
Const Sql$ = "Select * into [Sheet21] from [Sheet1$]"
FxSql_Arun Fx, Sql
End Sub

Private Sub FxSql_Drs__Tst()
Const Fx$ = "N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Const Sql$ = "Select * from [Sheet1$]"
DrsBrw FxSql_Drs(Fx, Sql)
End Sub

Sub FxWsNy__Tst()
AyDmp FxWsNy(SampleFx_KE24)
End Sub

Sub FxWs_Dt__Tst()
DtBrw FxWs_Dt(SampleFx_KE24)
End Sub

Sub FxWs_Fny__Tst()
AyDmp FxWs_Fny(SampleFx_KE24)
End Sub
