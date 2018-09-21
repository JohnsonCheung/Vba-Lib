Attribute VB_Name = "M_DbDt"
Option Explicit
Function DbDt_SqlAy_OfIns(A As Database, Dt As Dt) As String()
If DtIsEmp(Dt) Then Exit Function
Dim SimTyAy() As eSimTy
SimTyAy = DbtSimTyAy(A, Dt.DtNm)
Dim ValTp$
   ValTp = SimTyAy_InsValTp(SimTyAy)
Dim Tp$
   Dim T$, F$
   T = Dt.DtNm
   F = JnComma(Dt.Fny)
   Tp = FmtQQ("Insert into [?] (?) values(?)", T, F, ValTp)
Dim O$()
   Dim Dr
   ReDim O(UB(Dt.Dry))
   Dim J%
   J = 0
   For Each Dr In Dt.Dry
       O(J) = FmtQQAv(Tp, Dr)
       J = J + 1
   Next
DbDt_SqlAy_OfIns = O
End Function
Private Sub ZZ_DbDt_SqlAy_OfIns()
'Tmp1Tbl_Ens
Stop
Dim Db As Database: Set Db = Sample_Db_DutyPrepare
Dim Dt As Dt: Dt = DbtDt(Db, "Tmp1")
Dim O$(): O = DbDt_SqlAy_OfIns(Db, Dt)
Stop
End Sub
