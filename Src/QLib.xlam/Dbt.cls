VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Dbt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A As DAO.Database
Private T$

Property Get IsExist() As Property
IsExist = A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Property

Sub AddFld(F, Ty As DataTypeEnum)
Dim FF As New DAO.Field
FF.Name = F
FF.Type = Ty
Flds.Append FF
End Sub

Sub Brw()
DtBrw Dt
End Sub

Function Des$()
Des = PrpVal(A.TableDefs(T).Properties, "Description")
End Function

Function DftFny(Optional Fny0) As String()
If IsMissing(Fny0) Then
   DftFny = Me.Fny
Else
   DftFny = DftNy(Fny0)
End If
End Function

Sub Drp()
If IsExist Then A.Execute FmtQQ("Drop Table [?]", T)
End Sub

Function Dt() As Dt
Dim O As Dt
O.DtNm = T
O.Dry = RsDry(A.TableDefs(T).OpenRecordset)
O.Fny = Fny
Dt = O
End Function

Function Exist() As Boolean
Exist = Not A.OpenRecordset("Select Name from MSysObjects where Type in (1,6) and Name='?'").EOF
End Function

Function Flds() As DAO.Fields
Set Flds = A.TableDefs(T).Fields
End Function

Function Fny() As String()
Fny = FldsFny(A.TableDefs(T).Fields)
End Function

Function FxOfLnkTbl$()
FxOfLnkTbl = TakBet(A.TableDefs(T).Connect, "Database=", ";")
End Function

Function HasFld(F) As Boolean
Ass Dbt(A, T).IsExist
HasFld = TblHasFld(A.TableDefs(T), F)
End Function

Friend Property Get Init(Db As Database, TblNm) As Dbt
Set A = Db
T = TblNm
Set Init = Me
End Property

Sub LnkFb(Fb$, Optional SrcT0$)
Dim Src$: Src = Dft(SrcT0, T)
Dim Tbl  As TableDef
Set Tbl = A.CreateTableDef(T)
Tbl.SourceTableName = Src
Tbl.Connect = ";DATABASE=?" & Fb
Drp
A.TableDefs.Append Tbl
End Sub

Sub LnkFxWs(Fx$, Optional WsNm0)
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

Function Pk() As String()
Dim I  As Index, O$(), F
On Error GoTo X
If A.TableDefs(T).Indexes.Count = 0 Then Exit Function
On Error GoTo 0
For Each I In A.TableDefs(T).Indexes
   If I.Primary Then
       For Each F In I.Fields
           Push O, F.Name
       Next
       Pk = O
       Exit Function
   End If
Next
X:
End Function

Function RecCnt&()
RecCnt = DbqV(A, FmtQQ("Select Count(*) from [?]", T))
End Function

Function SimTyAy(Optional Fny0) As eSimTy()
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
End Function

Function StruLin$(Optional SkipTn As Boolean)
Dim O$(): O = Me.Fny: If AyIsEmp(O) Then Exit Function
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
End Function

Function TblFInfDry() As Variant()
Dim O(), F, Dr(), Fny$()
Fny = Me.Fny
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
TblFInfDry = O
End Function

Friend Sub Tst()
Pk__Tst
End Sub

Function Ws() As Worksheet
Set Ws = DtWs(Me.Dt)
End Function

Private Sub Pk__Tst()
Set A = SampleDb_DutyPrepare
Dim Dr(), Dry(), T
For Each T In DbTny(A)
    Erase Dr
    Push Dr, T
    PushAy Dr, DaoX.Dbt(A, CStr(T)).Pk
    Push Dry, Dr
Next
DryBrw Dry
End Sub
