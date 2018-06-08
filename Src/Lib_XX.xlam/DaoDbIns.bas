Attribute VB_Name = "DaoDbIns"
'Option Explicit
'
'Function DbDs_SqlAy_OfIns(A As Database, Ds As Ds) As String()
'If DsIsEmp(Ds) Then Exit Function
'Dim O$()
'Dim J%
'For J = 0 To UBound(Ds.DtAy)
'   PushAy O, DbDt_SqlAy_OfIns(A, Ds.DtAy(J))
'Next
'DbDs_SqlAy_OfIns = O
'End Function
'Function DbDt_SqlAy_OfIns(A As Database, Dt As Dt) As String()
'If DtIsEmp(Dt) Then Exit Function
'Dim SimTyAy() As eSimTy
'SimTyAy = Dbt(A, Dt.DtNm).SimTyAy(Dt.Fny)
'Dim ValTp$
'   ValTp = SimTyAy_InsValTp(SimTyAy)
'Dim Tp$
'   Dim T$, F$
'   T = Dt.DtNm
'   F = JnComma(Dt.Fny)
'   Tp = FmtQQ("Insert into [?] (?) values(?)", T, F, ValTp)
'Dim O$()
'   Dim Dr
'   ReDim O(UB(Dt.Dry))
'   Dim J%
'   J = 0
'   For Each Dr In Dt.Dry
'       O(J) = FmtQQAv(Tp, Dr)
'       J = J + 1
'   Next
'DbDt_SqlAy_OfIns = O
'End Function
'
'Sub DsInsDb(A As Ds, Db As Database)
'DbSqlAy_Run Db, DbDs_SqlAy_OfIns(Db, A)
'End Sub
'
'Sub DtInsDb(A As Database, Dt As Dt)
'DbSqlAy_Run A, DbDt_SqlAy_OfIns(A, Dt)
'End Sub
'
'Function SimTyAy_InsValTp$(SimTyAy() As eSimTy)
'Dim U%
'   U = UB(SimTyAy)
'Dim Ay$()
'   ReDim Ay(U)
'Dim J%
'For J = 0 To U
'   Ay(J) = SimTy_QuoteTp(SimTyAy(J))
'Next
'SimTyAy_InsValTp = JnComma(Ay)
'End Function
'
'Private Sub DbDt_SqlAy_OfIns__Tst()
''Tmp1Tbl_Ens
'Stop
'Dim Dt As Dt: Dt = Dbt(CurDb, "Tmp1").Dt
'Dim O$(): O = DbDt_SqlAy_OfIns(CurDb, Dt)
'Stop
'End Sub
'
