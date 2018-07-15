Attribute VB_Name = "M_Xls"
Option Explicit

Property Get NewWb(Optional Vis As Boolean) As Workbook
If Vis Then Xls.Visible = True
Set NewWb = Xls.Workbooks.Add
End Property

Property Get NewWs(Optional WsNm$ = "Sheet1", Optional Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWb(Vis).Sheets(1)
If O.Name <> WsNm Then O.Name = WsNm
Set NewWs = O
End Property

Property Get Xls() As Excel.Application
Static Y As New Excel.Application
On Error GoTo X
If Y.Name = "Microsoft Excel" Then
End If
E:
Set Xls = Y
Exit Property
X:
Set Y = New Excel.Application
GoTo E
End Property
