Attribute VB_Name = "M_Fb"
Option Explicit
Function FbDb(A) As Database
Set FbDb = DBEngine.OpenDatabase(A)
End Function
