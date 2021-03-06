VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Fx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Fx$

Property Get ACnStr$()
ACnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 12.0;HDR=YES""", Fx)
End Property

Property Get Cn() As ADODB.Connection
Dim O As New ADODB.Connection
O.Open ACnStr
Set Cn = O
End Property

Property Get FstWsNm$()
Dim T As ADOX.Table
Dim O$()
For Each T In Cat.Tables
    FstWsNm = RmvLasChr(T.Name)
    Exit Property
Next
End Property

Property Get WsNy() As String()
WsNy = SyRmvLasChr(ItrNy(Cat.Tables))
End Property

Function Cat() As Catalog
Dim O As New Catalog
Set O.ActiveConnection = Me.Cn
Set Cat = O
End Function

Function DftWsNy(WsNy0) As String()
Dim WsNy$(): WsNy = DftNy(WsNy0)
If AyIsEmp(WsNy) Then
   DftWsNy = Me.WsNy
   Exit Function
End If
DftWsNy = WsNy
End Function

Function HasWs(WsNm) As Boolean
HasWs = AyHas(WsNy, WsNm)
End Function

Sub RmvWsIfExist(WsNm)
If HasWs(WsNm) Then
   Dim B As Workbook: Set B = Wb
   WbWs(B, WsNm).Delete
   WbSav B
   WbClsNoSav B
End If
End Sub

Function SqlDrs(Sql) As Drs
SqlDrs = AdoP.Rs(Cn.Execute(Sql)).Drs
End Function

Sub SqlRun(Sql)
Cn.Execute CStr(Sql)
End Sub

Function Wb(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = NewXls.Workbooks.Open(Fx)
If Vis Then WbVis O
Set Wb = O
End Function

Function WsDt(Optional WsNm0$) As Dt
Dim WsNm$
If WsNm0 = "" Then WsNm = FstWsNm Else WsNm = WsNm0
Dim Sql$: Sql = FmtQQ("Select * from [?$]", WsNm)
WsDt = DtNmDrs_Dt(WsNm, AdoP.Cn(Cn).SqlDrs(Sql))
End Function

Function WsFny(Optional WsNm0$) As String()
Dim WsNm$: WsNm = DftWsNmByFxFstWs(WsNm, Fx)
WsFny = ItrNy(Cat.Tables(WsNm & "$").Columns)
End Function

Private Sub AqlDrs__Tst()
Const Sql$ = "Select * from [Sheet1$]"
DrsBrw Xls.Fx(SampleFx_KE24).SqlDrs(Sql)
End Sub

Private Sub Cat__Tst()
Dim A As ADOX.Catalog
Set A = Xls.Fx(SampleFx_KE24).Cat
Stop
End Sub

Private Sub Cn__Tst()
Dim A As ADODB.Connection
Set A = Xls.Fx(SampleFx_KE24).Cn
End Sub

Private Sub RmvWsIfExist__Tst()
Me.Fx = TmpFx("FxRmvWsIfExist__Tst")

Dim Wb As Workbook
Set Wb = NewWb
Wb.Sheets.Add
Wb.SaveAs Fx
Dim WsNyBef$(), WsNyAft$()
   WsNyBef = WsNy
   RmvWsIfExist "Sheet1"
   WsNyAft = WsNy
Dim Exp$()
   Exp = AyMinus(WsNyBef, Array("Sheet1"))
AyPair_EqChk Exp, WsNyAft
End Sub

Private Sub SqlRun__Tst()
Const Fx$ = SampleFx_KE24
Const Sql$ = "Select * into [Sheet21] from [Sheet1$]"
Xls.Fx(Fx).SqlRun Sql
End Sub

Private Sub Sql_Drs__Tst()
Const Fx$ = "N:\SapAccessReports\DutyPrepay5\SAPDownloadExcel\KE24 2010-01c.xls"
Const Sql$ = "Select * from [Sheet1$]"
DrsBrw Xls.Fx(Fx).SqlDrs(Sql)
End Sub

Sub WsDt__Tst()
DtBrw Xls.Fx(SampleFx_KE24).WsDt
End Sub

Sub WsFny__Tst()
AyDmp Xls.Fx(SampleFx_KE24).WsFny
End Sub

Sub WsNy__Tst()
AyDmp Xls.Fx(SampleFx_KE24).WsNy
End Sub
