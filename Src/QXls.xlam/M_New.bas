Attribute VB_Name = "M_New"
Option Explicit

Function NewWb(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = NewXls.Workbooks.Add
If Vis Then O.Application.Visible = True
Set NewWb = O
End Function

Function NewWs(Optional WsNm$, Optional Vis As Boolean) As Worksheet
Dim Wb As Workbook
Set Wb = NewWb
WsDlt Wb, "Sheet2"
WsDlt Wb, "Sheet3"
If WsNm <> "" Then WbWs(Wb, "Sheet1").Name = WsNm
Set NewWs = WbWs(Wb, 1)
If Vis Then WbVis Wb
End Function

Function NewWsA1() As Range
Set NewWsA1 = WsA1(NewWs)
End Function

Function NewXls() As Excel.Application
Static X As Excel.Application
On Error GoTo XX
Dim A$: A = X.Name
Set NewXls = X
Exit Function
XX:
Set X = New Excel.Application
Set NewXls = X
End Function

