Attribute VB_Name = "M_Cur"
Option Explicit
Function CurDb() As Database
Set CurDb = Acs.CurrentDb
End Function
Function CurDbPth$()
CurDbPth = FfnPth(CurFb)
End Function
Function CurFb$()
On Error Resume Next
CurFb = CurDb.Name
End Function
