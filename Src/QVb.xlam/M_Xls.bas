Attribute VB_Name = "M_Xls"
Option Explicit

Function Xls() As Excel.Application
Static Y As Excel.Application
On Error GoTo X
Dim A$: A = Y.Name
Set Xls = Y
Exit Function
X:
Set Y = New Excel.Application
Set Xls = Y
End Function
