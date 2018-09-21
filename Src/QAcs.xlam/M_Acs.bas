Attribute VB_Name = "M_Acs"
Option Explicit
Function Acs() As Access.Application
Static X As Access.Application
On Error GoTo XX
Dim A$: A = X.Name
Set Acs = X
Exit Function
XX:
Set X = New Access.Application
Set Acs = X
End Function

