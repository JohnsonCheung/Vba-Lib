Attribute VB_Name = "M_New"
Option Explicit

Function NewWb(Optional Vis As Boolean) As Workbook
If Vis Then Xls.Visible = True
Set NewWb = Xls.Workbooks.Add
End Function

Function NewWs(Optional WsNm$ = "Sheet1", Optional Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWb(Vis).Sheets(1)
If O.Name <> WsNm Then O.Name = WsNm
Set NewWs = O
End Function
