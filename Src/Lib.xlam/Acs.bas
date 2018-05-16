Attribute VB_Name = "Acs"
Option Explicit

Function CurAcs() As Access.Application
Static X As Access.Application
On Error GoTo XX
Dim A$: A = X.Name
Set CurAcs = X
Exit Function
XX:
Set X = New Access.Application
Set CurAcs = X
End Function

Function CurDb() As Database
Set CurDb = CurAcs.CurrentDb
End Function

Function CurDbPth$()
CurDbPth = FfnPth(CurFb)
End Function

Function CurFb$()
CurFb = CurrentDb.Name
End Function

Function WrkPth$()
WrkPth = CurDbPth & "WorkingDir\"
End Function

