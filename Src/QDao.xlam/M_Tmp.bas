Attribute VB_Name = "M_Tmp"
Option Explicit
Function TmpDb(Optional Fnn$) As Database
Set TmpDb = DBEngine.CreateDatabase(TmpFb("TmpDb", Fnn), DAO.LanguageConstants.dbLangGeneral)
End Function
