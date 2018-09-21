Attribute VB_Name = "M_Dft"
Option Explicit
Function CurDb() As Database
Stop '
End Function
Function DftDb(A As Database) As Database
If IsNothing(A) Then
   Set DftDb = CurDb
Else
   Set DftDb = A
End If
End Function
Function DftFb$(A$)
If A = "" Then
   Dim O$: O = TmpFb
   DAO.DBEngine.CreateDatabase(O, dbLangGeneral).Close
   DftFb = O
Else
   DftFb = A
End If
End Function
Function DftDbNm$(DbNm0$, Db As Database)
If DbNm0 = "" Then
    DftDbNm = FFnFnn(Db.Name)
Else
    DftDbNm = DbNm0
End If
End Function
