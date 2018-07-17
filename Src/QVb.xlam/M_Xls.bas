Attribute VB_Name = "M_Xls"
Option Explicit

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
