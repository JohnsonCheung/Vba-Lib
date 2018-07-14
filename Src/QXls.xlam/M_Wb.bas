Attribute VB_Name = "M_Wb"
Option Explicit

Property Get WbAddWs(A As Workbook, Optional WsNm$, Optional IsBeg As Boolean) As Worksheet
Dim O As Worksheet
If IsBeg Then
    Set O = A.Sheets.Add(WbFstWs(A))
Else
    Set O = A.Sheets.Add(, WbLasWs(A))
End If
If WsNm <> "" Then
   O.Name = WsNm
End If
Set WbAddWs = O
End Property

Property Get WbFstWs(A As Workbook) As Worksheet
Set WbFstWs = A.Sheets(1)
End Property

Property Get WbHasWs(A As Workbook, Ix_or_WsNm) As Boolean
On Error GoTo X
Dim Ws As Worksheet: Set Ws = A.Sheets(Ix_or_WsNm)
WbHasWs = True
Exit Property
X:
End Property
Sub WbClsNoSav(A As Workbook)
On Error Resume Next
A.Close False
End Sub

Sub WbSav(A As Workbook)
Dim X As Excel.Application
Set X = A.Application
Dim Y As Boolean
Y = X.DisplayAlerts
A.DisplayAlerts = False
A.Save
A.DisplayAlerts = Y
End Sub

Sub WbVis(A As Workbook)
A.Application.Visible = True
End Sub

Property Get WbLasWs(A As Workbook) As Worksheet
Set WbLasWs = A.Sheets(A.Sheets.Count)
End Property

Property Get WbWs(A As Workbook, Ix_or_WsNm) As Worksheet
Set WbWs = A.Sheets(Ix_or_WsNm)
End Property

Property Get WbWsNy(A As Workbook) As String()
Stop
'WbWsNy = ItrNy(A.Sheets)
End Property

