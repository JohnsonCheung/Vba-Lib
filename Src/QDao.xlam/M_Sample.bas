Attribute VB_Name = "M_Sample"
Option Explicit
Public Const Sample_Fb_DutyPrepare$ = "C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Data.mdb"
Public Const Sample_Fb_DutyPrepareBackup$ = "C:\Users\User\Desktop\SAPAccessReports\DutyPrepay5\DutyPrepay5_Backup.accdb"
Function Sample_Db_DutyPrepare() As Database
Set Sample_Db_DutyPrepare = DBEngine.OpenDatabase(Sample_Fb_DutyPrepare)
End Function
