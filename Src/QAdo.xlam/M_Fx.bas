Attribute VB_Name = "M_Fx"
Option Explicit
Private Sub ZZ_FxWsNy()
AyDmp FxWsNy(SampleFx_KE24)
End Sub

Private Sub ZZ_FxFstWsNm()
Debug.Print FxFstWsNm(SampleFx_KE24)
End Sub

Function FxFstWsNm$(A)
FxFstWsNm = RmvLasChr(ItrFstNm(FxCat(A).Tables))
End Function

Function FxCn(A) As Connection
Dim O As New Connection
O.Open FxCnStr(A)
Set FxCn = O
End Function

Function FxCnStr$(A)
FxCnStr = FmtQQ("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=?;Extended Properties=""Excel 12.0;HDR=YES""", A)
End Function

Function FxWsNy(A, Optional Patn$ = ".") As String()
Dim O$(), I, N$
If Patn = "." Then
    For Each I In FxCat(A).Tables
        N = ObjNm(I)
        If LasChr(N) = "$" Then
            Push O, N
        End If
    Next
Else
    Dim R As RegExp: Set R = Re(Patn)
    For Each I In FxCat(A).Tables
        N = ObjNm(I)
        If LasChr(N) = "$" Then
            If R.Test(N) Then
                Push O, N
            End If
        End If
    Next
End If
FxWsNy = AyRmvLasChr(O)
End Function

Function FxCat(A) As Catalog
Dim O As New Catalog
Set O.ActiveConnection = FxCn(A)
Set FxCat = O
End Function

Function FxDftWsNy(A, WsNy0) As String()
Dim WsNy$(): WsNy = DftNy(WsNy0)
If AyIsEmp(WsNy) Then
   FxDftWsNy = FxWsNy(A)
   Exit Function
End If
FxDftWsNy = WsNy
End Function

Function FxHasWs(A, WsNm) As Boolean
FxHasWs = CatHasTbl(FxCat(A), WsNm)
End Function

Sub FxRmvWsIfExist(A, WsNm)
If FxHasWs(A, WsNm) Then
   Dim B As Workbook: Set B = FxWb(A)
   WbWs(B, WsNm).Delete
   WbSav B
   WbClsNoSav B
End If
End Sub

Function FxSqlDrs(A, Sql) As Drs
Set FxSqlDrs = RsDrs(FxCn(A).Execute(Sql))
End Function

Sub FxSqlRun(A, Sql)
FxCn(A).Execute Sql
End Sub

Function FxWb(A, Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = Xls.Workbooks.Open(A)
If Vis Then WbVis O
Set FxWb = O
End Function

Function FxWsDt(A, Optional WsNm0$) As Dt
Dim N$: N = FxDftWsNm(A, WsNm0)
Dim Sql$: Sql = FmtQQ("Select * from [?$]", N)
Set FxWsDt = DrsDt(FxSqlDrs(A, Sql), N)
End Function

Function FxDftWsNm$(A, WsNm0$)
If WsNm0 = "" Then
    FxDftWsNm = FxFstWsNm(A)
    Exit Function
End If
FxDftWsNm = WsNm0
End Function

Function FxWsFny(A, Optional WsNm0$) As String()
Dim WsNm$: WsNm = DftWsNmByFxFstWs(WsNm, A)
FxWsFny = ItrNy(FxCat(A).Tables(WsNm & "$").Columns)
End Function
