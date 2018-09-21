Attribute VB_Name = "JDao"
Option Explicit

Function FbDb(A) As Database
Set FbDb = DAO.DBEngine.OpenDatabase(A)
End Function
