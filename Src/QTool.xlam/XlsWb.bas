Attribute VB_Name = "XlsWb"
Option Explicit
Function WbAddWs(A As Workbook, Optional WsNm$ = "Sheet1") As Worksheet
Dim O As Worksheet
Set O = A.Sheets.Add(A.Sheets(1))
If WsNm <> "" Then
   O.Name = WsNm
End If
Set WbAddWs = O
End Function
Function WbCn_TxtCn(A As WorkbookConnection) As TextConnection
On Error Resume Next
Set WbCn_TxtCn = A.TextConnection
End Function
Function WbTxtCn(A As Workbook) As TextConnection
Dim N%: N = WbTxtCnCnt(A)
If N <> 1 Then
    Stop
    Exit Function
End If
Dim C As WorkbookConnection
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then
        Set WbTxtCn = C.TextConnection
        Exit Function
    End If
Next
ErImposs
End Function
Function WbTxtCnCnt%(A As Workbook)
Dim C As WorkbookConnection, Cnt%
For Each C In A.Connections
    If Not IsNothing(WbCn_TxtCn(C)) Then Cnt = Cnt + 1
Next
WbTxtCnCnt = Cnt
End Function
Function WbTxtCnStr$(A As Workbook)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
If IsNothing(T) Then Exit Function
WbTxtCnStr = T.Connection
End Function
Function WbVis(A As Workbook) As Workbook
XlsVis A.Application
Set WbVis = A
End Function
Function WbSavAs(A As Workbook, Fx) As Workbook
A.SaveAs Fx
Set WbSavAs = A
End Function
Function WbRfh(A As Workbook) As Workbook
Dim Ws As Worksheet
For Each Ws In A.Worksheets
    WsRfh Ws
Next
Dim PC As PivotCache
For Each PC In A.PivotCaches
    PC.MissingItemsLimit = xlMissingItemsNone
    PC.Refresh
Next
Set WbRfh = A
End Function
Sub WbSetFcsv(A As Workbook, Fcsv$)
'Assume there is one and only one TextConnection.  Set it using {Fcsv}
Dim T As TextConnection: Set T = WbTxtCn(A)
Dim C$: C = T.Connection: If Not HasPfx(C, "TEXT;") Then Stop
T.Connection = "TEXT;" & Fcsv
End Sub
Private Sub Z_WbSetFcsv()
Dim Wb As Workbook
Set Wb = FxWb(VbeMthFx)
Debug.Print WbTxtCnStr(Wb)
WbSetFcsv Wb, "C:\ABC.CSV"
Ass WbTxtCnStr(Wb) = "TEXT;C:\ABC.CSV"
Wb.Close False
Stop
End Sub
Private Sub Z_WbTxtCnCnt()
Dim O As Workbook: Set O = FxWb(VbeMthFx)
Ass WbTxtCnCnt(O) = 1
O.Close
End Sub
