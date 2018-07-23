Attribute VB_Name = "M_Dft"
Option Explicit
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
   Dao.DBEngine.CreateDatabase(O, dbLangGeneral).Close
   DftFb = O
Else
   DftFb = A
End If
End Function
