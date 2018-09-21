Attribute VB_Name = "M_Dbt"
Option Explicit

Function DbtDes$(A As Database, T)
DbtDes = PrpVal(A.TableDefs(T).Properties, "Description")
End Function

Function DbtDftFny(A As Database, T, Optional Fny0) As String()
If IsMissing(Fny0) Then
   DbtDftFny = DbtFny(A, T)
Else
   DbtDftFny = DftNy(Fny0)
End If
End Function

Function DbtDt(A As Database, T) As Dt
Dim Fny$(): Fny = DbtFny(A, T)
Dim Dry(): Dry = RsDry(DbtRs(A, T))
Set DbtDt = Dt(T, Fny, Dry)
End Function

Function DbtRs(A As Database, T) As Recordset
Set DbtRs = A.TableDefs(T).OpenRecordset
End Function
Function DbtIsExist(A As Database, T) As Boolean
DbtIsExist = Not A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Function DbtFlds(A As Database, T) As DAO.Fields
Set DbtFlds = A.TableDefs(T).Fields
End Function

Function DbtFny(A As Database, T) As String()
DbtFny = FldsFny(A.TableDefs(T).Fields)
End Function

Function DbtFxOfLnkTbl$(A As Database, T)
DbtFxOfLnkTbl = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Function DbtHasFld(A As Database, T, F) As Boolean
Ass DbtIsExist(A, T)
DbtHasFld = TblHasFld(A.TableDefs(T), F)
End Function

Function DbtPk(A As Database, T) As String()
Dim I  As Index, O$(), F
On Error GoTo X
If A.TableDefs(T).Indexes.Count = 0 Then Exit Function
On Error GoTo 0
For Each I In A.TableDefs(T).Indexes
   If I.Primary Then
       For Each F In I.Fields
           Push O, F.Name
       Next
       DbtPk = O
       Exit Function
   End If
Next
X:
End Function

Function DbtRecCnt&(A As Database, T)
DbtRecCnt = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
End Function

Function DbtSimTyAy(A As Database, T, Optional Fny0) As eSimTy()
Dim Fny$(): Fny = DftNy(Fny0)
Dim O() As eSimTy
   Dim U%
   ReDim O(U)
   Dim J%, F
   J = 0
   For Each F In Fny
       O(J) = DaoTy_SimTy(DbTF_Fld(A, T, F).Type)
       J = J + 1
   Next
DbtSimTyAy = O
End Function

Function DbtStruLin$(A As Database, T, Optional SkipTn As Boolean)
Dim O$(): O = DbtFny(A, T): If AyIsEmp(O) Then Exit Function
O = FnyQuote(O, DbtPk(A, T))
O = FnyQuoteIfNeed(O)
Dim J%, V
V = 0
For Each V In O
   O(J) = Replace(V, T, "*")
   J = J + 1
Next
If SkipTn Then
   DbtStruLin = JnSpc(O)
Else
   DbtStruLin = T & " = " & JnSpc(O)
End If
End Function

Function DbtInfDryzTblF(A As Database, T) As Variant()
Dim O(), F, Dr(), Fny$()
Fny = DbtFny(A, T)
If AyIsEmp(Fny) Then Exit Function
Dim SeqNo%
SeqNo = 0
For Each F In Fny
    Erase Dr
    Push Dr, T
    Push Dr, SeqNo: SeqNo = SeqNo + 1
    PushAy Dr, DbTF_FldInfDr(A, T, CStr(F))
    Push O, Dr
Next
DbtInfDryzTblF = O
End Function

Function DbtWs(A As Database, T) As Worksheet
Set DbtWs = DtWs(DbtDt(A, T))
End Function

Sub DbtAddFld(A As Database, T, F, Ty As DataTypeEnum, Optional Sz%, Optional Precious%)
If DbtHasFld(A, T, F) Then Exit Sub
Dim S$, SqlTy$
SqlTy = DaoTy_SqlTy(Ty, Sz, Precious)
S = FmtQQ("Alter Table [?] Add Column [?] ?", T, F, Ty)
A.Execute S
End Sub
Function DaoTy_SqlTy$(A As DataTypeEnum, Optional Sz%, Optional Precious%)
Stop '
End Function

Sub DbtBrw(A As Database, T)
DtBrw DbtDt(A, T)
End Sub

Sub DbtDrp(A As Database, T)
If DbtIsExist(A, T) Then A.Execute FmtQQ("Drop Table [?]", T)
End Sub

Sub DbtLnkFb(A As Database, T, Fb$, Optional SrcT0$)
Dim Src$: Src = Dft(SrcT0, T)
Dim Tbl  As TableDef
Set Tbl = A.CreateTableDef(T)
Tbl.SourceTableName = Src
Tbl.Connect = ";DATABASE=?" & Fb
A.TableDefs.Append Tbl
End Sub
Function DftWsNm$(WsNm0$, Fx$)
If WsNm0 = "" Then
    DftWsNm = WsNm0
Else
    DftWsNm = FxFstWsNm(Fx)
End If
End Function
Function FxFstWsNm$(Fx$)
Stop '
End Function
Sub DbtLnkFxWs(A As Database, T, Fx$, Optional WsNm0$)
Const CSub$ = "ATLnkFxWs"
Dim WsNm$: WsNm = DftWsNm(WsNm0, Fx)
On Error GoTo X
   Dim Tbl  As TableDef
   Set Tbl = A.CreateTableDef(T)
   Tbl.SourceTableName = WsNm & "$"
   Tbl.Connect = FmtQQ("Excel 8.0;HDR=YES;IMEX=2;DATABASE=?", Fx)
   DbtDrp A, T
   A.TableDefs.Append Tbl
Exit Sub
X: Er CSub, "{Er} found in Creating {T} in {Db} by Linking {WsNm} in {Fx}", Err.Description, T, A.Name, WsNm0, Fx
End Sub

Private Sub ZZ_DbtPk()
Dim A As Database
Set A = Sample_Db_DutyPrepare
Dim Dr(), Dry(), T
For Each T In DbTny(A)
    Erase Dr
    Push Dr, T
    PushAy Dr, DbtPk(A, T)
    Push Dry, Dr
Next
DryBrw Dry
End Sub
