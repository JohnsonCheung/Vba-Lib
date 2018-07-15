Attribute VB_Name = "JDao"
Option Explicit

Property Get Db() As Database
Set Db = Dao.DBEngine.OpenDatabase(A)
End Property
