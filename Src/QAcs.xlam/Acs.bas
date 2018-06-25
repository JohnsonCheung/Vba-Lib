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
CurDbPth = File(CurFb).Pth
End Function

Function CurFb$()
On Error Resume Next
CurFb = CurrentDb.Name
End Function

Function WrkPth$()
WrkPth = File(Excel.Application.VBE.ActiveVBProject.Filename).Pth & "WorkingDir\"
End Function
