Attribute VB_Name = "M_Dbt"
Option Explicit

Property Get DbtStruLin$(A As Database, T, Optional SkipTn As Boolean)
Dim O$(): O = Me.Fny: If AyIsEmp(O) Then Exit Property
O = FnyQuote(O, Me.Pk)
O = FnyQuoteIfNeed(O)
Dim J%, V
V = 0
For Each V In O
   O(J) = Replace(V, T, "*")
   J = J + 1
Next
If SkipTn Then
   StruLin = JnSpc(O)
Else
   StruLin = T & " = " & JnSpc(O)
End If
End Property

Property Get DbtExist(A As Database, T) As Boolean
Exist = Not A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Property

Property Get DbtFlds(A As Database, T) As Dao.Fields
Set Flds = A.TableDefs(T).Fields
End Property

Property Get DbtFny(A As Database, T) As String()
Fny = FldsFny(A.TableDefs(T).Fields)
End Property

Property Get DbtFxOfLnkTbl$(A As Database, T)
FxOfLnkTbl = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Property

Sub DbtLnkFb(A As Database, T, Fb$, Optional SrcT0$)
Dim Src$: Src = Dft(SrcT0, T)
Dim Tbl  As TableDef
Set Tbl = A.CreateTableDef(T)
Tbl.SourceTableName = Src
Tbl.Connect = ";DATABASE=?" & Fb
Drp
A.TableDefs.Append Tbl
End Sub
Sub DbtBrw(A As Database, T)
DtBrw Dt
End Sub

Property Get DbtDes$(A As Database, T)
Des = PrpVal(A.TableDefs(T).Properties, "Description")
End Property

Sub DbtDrp(A As Database, T)
If DbtIsExist(A, T) Then A.Execute FmtQQ("Drop Table [?]", T)
End Sub
Private Sub ZZ_DbtPk()
Set A = SampleDb_DutyPrepare
Dim Dr(), Dry(), T
For Each T In DbTny(A)
    Erase Dr
    Push Dr, T
    PushAy Dr, DbtPk(A, T)
    Push Dry, Dr
Next
DryBrw Dry
End Sub
Property Get DbtHasFld(A As Database, T, F) As Boolean
Ass DbtIsExist(A, T)
DbtHasFld = TblHasFld(A.TableDefs(T), F)
End Property

Property Get DbtDt(A As Database, T) As Dt
Set DbtDt = Dt(T, DbtFny(A, T), RsDry(A.TableDefs(T).OpenRecordset))
End Property

Sub DbtLnkFxWs(Fx$, Optional WsNm0)
Const CSub$ = "ATLnkFxWs"
Dim WsNm$: WsNm = Dft(WsNm0, T)
On Error GoTo X
   Dim Tbl  As TableDef
   Set Tbl = A.CreateTableDef(T)
   Tbl.SourceTableName = WsNm & "$"
   Tbl.Connect = FmtQQ("Excel 8.0;HDR=YES;IMEX=2;DATABASE=?", Fx)
   Drp
   A.TableDefs.Append Tbl
Exit Sub
X: Er CSub, "{Er} found in Creating {T} in {Db} by Linking {WsNm} in {Fx}", Err.Description, T, A.Name, WsNm0, Fx
End Sub

Property Get DbtPk(A As Database, T) As String()
Dim I  As Index, O$(), F
On Error GoTo X
If A.TableDefs(T).Indexes.Count = 0 Then Exit Property
On Error GoTo 0
For Each I In A.TableDefs(T).Indexes
   If I.Primary Then
       For Each F In I.Fields
           Push O, F.Name
       Next
       Pk = O
       Exit Property
   End If
Next
X:
End Property

Property Get DbtTblFInfDry() As Variant()
Dim O(), F, Dr(), Fny$()
Fny = Me.Fny
If AyIsEmp(Fny) Then Exit Property
Dim SeqNo%
SeqNo = 0
For Each F In Fny
    Erase Dr
    Push Dr, T
    Push Dr, SeqNo: SeqNo = SeqNo + 1
    PushAy Dr, DbTF_FldInfDr(A, T, CStr(F))
    Push O, Dr
Next
TblFInfDry = O
End Property

Property Get DbtRecCnt&()
RecCnt = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
End Property
Property Get DbtDftFny(Optional Fny0) As String()
If IsMissing(Fny0) Then
   DftFny = Me.Fny
Else
   DftFny = DftNy(Fny0)
End If
End Property


Sub DbtAddFld(F, Ty As DataTypeEnum)
Dim FF As New Dao.Field
FF.Name = F
FF.Type = Ty
Flds.Append FF
End Sub

Property Get DbtSimTyAy(Optional Fny0) As eSimTy()
Dim Fny$(): Fny = DftFny(Fny0)
Dim O() As eSimTy
   Dim U%
   ReDim O(U)
   Dim J%, F
   J = 0
   For Each F In Fny
       O(J) = NewSimTy(DbTF_Fld(A, T, F).Type)
       J = J + 1
   Next
SimTyAy = O
End Property

Property Get DbtWs() As Worksheet
Set Ws = DtWs(Me.Dt)
End Property


