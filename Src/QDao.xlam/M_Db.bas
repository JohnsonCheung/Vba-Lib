Attribute VB_Name = "M_Db"
Option Explicit
Property Get DbDs(A As Database, Tny0, Optional DsNm$ = "Ds") As Ds
Dim DtAy1() As Dt
    Dim U%, Tny$()
    Tny = DftNy(Tny0)
    U = UB(Tny)
    ReDim DtAy(U)
    Dim J%
    For J = 0 To U
        DtAy(J) = DbtDt(A, Tny(J))
    Next
Set DbDs = Ds(DsNm, DtAy1)
End Property

Private Sub ZZ_DbDs()
Dim Ds As Ds
Set Ds = DbDs(CurDb, "Permit PermitD")
Stop
End Sub
Sub DbCrtTbl(A As Database, T, FldDclAy)
A.Execute FmtQQ("Create Table [?] (?)", T, JnComma(FldDclAy))
End Sub

Sub DbEnsTmp1Tbl(A As Database)
If DbHasTbl(A, "Tmp1") Then Exit Sub
DbqRun A, "Create Table Tmp1 (AA Int, BB Text 10)"
End Sub

Function DbHasTbl(A As Database, T) As Boolean
DbHasTbl = A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Sub DbLnkFb(A As Database, Fb$, Tny0, Optional SrcTny0)
Dim Tny$(): Tny = DftNy(Tny0)              ' Src_Tny
Dim Src$(): Src = DftNy(Dft(SrcTny0, Tny0)) ' Tar_Tny
Ass Sz(Tny) > 0
Ass AyPair_IsEqSz(Src, Tny)
Dim J%
For J = 0 To UB(Tny)
    Dbt(A, Tny(J)).LnkFb Fb, Src(J)
Next
End Sub

Sub DbLnkFx(A As Database, Fx$, Optional WsNy0)
Dim WsNy$(): WsNy = Xls.Fx(Fx).DftWsNy(WsNy0)
Dim J%
For J = 0 To UB(WsNy)
   Dbt(A, WsNy(J)).LnkFxWs Fx
Next
End Sub

Function DbQny(A As Database) As String()
DbQny = DbqSy(A, "Select Name from MSysObjects where Type=5 and Left(Name,4)<>'MSYS' and Left(Name,4)<>'~sq_'")
End Function

Property Get Ds(Tny0) As Ds
Dim Tny$(): Tny = DftNy(Tny)
End Property

Private Sub ZZ_Qny()
AyDmp DbQny(CurDb)
End Sub

Property Get Db_DbInf(A As Database) As DbInf
Stop '
End Property


