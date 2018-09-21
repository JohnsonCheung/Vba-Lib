Attribute VB_Name = "M_PgmDb"
Option Explicit
Function PgmDb_DtaDb(A As Database) As Database
Set PgmDb_DtaDb = DBEngine.OpenDatabase(PgmDb_DtaFb(A))
End Function
Function PgmDb_DtaFb$(A As Database)

End Function
