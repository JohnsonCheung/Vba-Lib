Attribute VB_Name = "XlsWs"
Option Explicit
Function WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Function
Function WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Function
Function WsA1(A As Worksheet) As Range
Set WsA1 = A.Cells(1, 1)
End Function
Function WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Function
Function WsRR(A As Worksheet, R1, R2) As Range
Set WsRR = A.Range(WsRC(A, R1, 1), WsRC(A, R2, 1)).EntireRow
End Function
Function WsVis(A As Worksheet) As Worksheet
XlsVis A.Application
Set WsVis = A
End Function
Sub WsRfh(A As Worksheet)
Dim L As ListObject, Qt As QueryTable
For Each L In A.ListObjects
    Set Qt = LoQt(L)
    If Not IsNothing(Qt) Then Qt.Refresh False
Next
Dim Q As QueryTable
For Each Q In A.QueryTables
    Q.Refresh False
Next
Dim P As PivotTable
For Each P In A.PivotTables
    P.RefreshTable
Next
End Sub
