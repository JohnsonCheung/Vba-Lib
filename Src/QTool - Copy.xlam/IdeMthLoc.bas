Attribute VB_Name = "IdeMthLoc"
Option Explicit
Public Const MthFb$ = "C:\Users\User\Desktop\Vba-Lib-1\Mth.accdb"
Public Const MthLocFx$ = "C:\Users\User\Desktop\Vba-Lib-1\MthLoc.xlsx"
Public Const WrkFb$ = "C:\Users\User\Desktop\Vba-Lib-1\MthWrk.accdb"
Private Q$
Sub BrwMthFb()
FbBrw MthFb
End Sub
Function W() As Database
Static Y As Database, X As Boolean
If Not X Then
    X = False
    FbEns WrkFb
    Set Y = FbDb(WrkFb)
End If
Set W = Y
End Function
Sub DbttLnkFbtt(A As Database, TT, Fb$, Optional Fbtt0)
Dim Tny$(): Tny = CvNy(TT)
Dim Fbtt$(): Fbtt = CvNy(Fbtt)
If Sz(Fbtt) = 0 Then Fbtt = Tny
Dim J%, T
For Each T In Tny
    DbtLnkFbt A, T, Fb, Fbtt(J)
    J = J + 1
Next
End Sub
Function DbtLnk(A As Database, T, S$, Cn$) As String()
On Error GoTo X
Dim TT As New DAO.TableDef
DbttDrp A, T
With TT
    .Connect = Cn
    .Name = T
    .SourceTableName = S
    A.TableDefs.Append TT
End With
Exit Function
X:
Dim Er$
Er = Err.Description
Debug.Print Er
Dim O$(), M$
M = "Cannot create Table in Database from Source by Cn with Er from system"
PushI O, "Program  : DbtLnk"
PushI O, "Database : " & A.Name
PushI O, "Table    : " & T
PushI O, "Source   : " & S
PushI O, "Cn       : " & Cn
PushI O, "Er       : " & Er
PushI O, M
DbtLnk = O
End Function
Function DftStr$(A, Dft)
DftStr = IIf(A = "", Dft, A)
End Function

Function DbtLnkFbt(A As Database, T, Fb$, Fbt0$) As String()
Dim Fbt$, Cn$
Cn = ";Database=" & Fb
Fbt = DftStr(Fbt0, T)
DbtLnkFbt = DbtLnk(A, T, Fbt, Cn)
End Function

Sub WBrw()
FbBrw WrkFb
End Sub

Sub Gen()
DbttLnkFbtt W, "Mth MthLoc", MthFb
GenCrtDistLines
GenRmvDistLinesDclOpt
GenCrtDistMth
GenUpdMthLoc
GenCrtMdDic
GenCrtClsDic
End Sub
Sub GenCrtClsDic()
DbttDrp W, "ClsDic #A"
Q = "Select Md,Lines into [#A] from Mth where MdTy='Cls' order by Md,Nm": W.Execute Q
DbtCrtFldLisTbl W, "#A", "ClsDic", "Md", "Lines", vbCrLf & vbCrLf, True, "Lines"
End Sub
Function DclRmvOpt$(A)
Dim Ly$(): Ly = SplitCrLf(A)
Dim L, O$()
For Each L In AyNz(Ly)
    If Not HasPfx(L, "Option ") Then
        PushI O, L
    End If
Next
DclRmvOpt = JnCrLf(O)
End Function
Sub GenRmvDistLinesDclOpt()
With DbqRs(W, "Select Lines from DistLines where Nm='*Dcl'")
    While Not .EOF
        .Edit
        .Fields(0).Value = DclRmvOpt(.Fields(0).Value)
        .Update
        .MoveNext
    Wend
End With
End Sub

Sub GenCrtDistLines()
DbttDrp W, "DistLines #A"
Q = "Select Nm,Lines into [#A] from Mth where MdTy='Std'": W.Execute Q
Q = "Alter Table [#A] Add Column LinesId Long": W.Execute Q
DbtUpdIdFld W, "#A", "Lines"
Q = "Select Distinct Nm,LinesId Into DistLines from [#A]": W.Execute Q
Q = "Alter Table DistLines Add Column Lines Memo": W.Execute Q
Dim D As New Dictionary, V, Ix&
With DbqRs(W, "Select LinesId,Lines from [#A]")
    While Not .EOF
        V = .Fields(1).Value
        If Not D.Exists(V) Then
            D.Add V, Ix
            Ix = Ix + 1
        End If
        .MoveNext
    Wend
    .Close
End With
Dim Ay$(), K
ReDim Ay$(D.Count)
For Each K In D.Keys
    Ay(D(K)) = K
Next

With DbqRs(W, "Select LinesId,Lines from [DistLines]")
    While Not .EOF
        .Edit
        !Lines = Ay(.Fields(0).Value)
        .Update
        .MoveNext
    Wend
    .Close
End With
End Sub
Sub GenCrtDistMth()
DbttDrp W, "DistMth #A #B"
Q = "Select Distinct Nm,Count(*) as LinesIdCnt Into DistMth from DistLines group by Nm": W.Execute Q
Q = "Alter Table DistMth Add Column LinesIdLis Text(255), LinesLis Memo, ToMd Text(50)": W.Execute Q
DbtCrtFldLisTbl W, "DistLines", "#A", "Nm", "LinesId", " ", True
DbtCrtFldLisTbl W, "DistLines", "#B", "Nm", "Lines", vbCrLf & vbCrLf, True
Q = "Update DistMth x inner join [#A] a on x.Nm = a.Nm set x.LinesIdLis = a.LinesIdLis": W.Execute Q
Q = "Update DistMth x inner join [#B] a on x.Nm = a.Nm set x.LinesLis = a.LinesLis": W.Execute Q
Q = "Update DistMth x inner join MthLoc a on x.Nm = a.Nm set x.ToMd = IIf(a.ToMd='','AAMod',a.ToMd)": W.Execute Q
End Sub
Sub GenUpdMthLoc()
DbtDrp W, "#A"
Q = "Select x.Nm into [#A] from DistMth x left join MthLoc a on x.Nm=a.Nm where IsNull(a.Nm)": W.Execute Q
Q = "Insert into MthLoc (Nm) Select Nm from [#A]": W.Execute Q
End Sub
Sub GenCrtMdDic()
DbtDrp W, "MdDic"
DbtCrtFldLisTbl W, "DistMth", "MdDic", "ToMd", "LinesLis", vbCrLf & vbCrLf, True, "Lines"
End Sub
Function DbqRs(A As Database, Q) As DAO.Recordset
Set DbqRs = A.OpenRecordset(Q)
End Function
Function RsAny(A As DAO.Recordset) As Boolean
If A.EOF Then Exit Function
If A.BOF Then Exit Function
RsAny = True
End Function
Sub DbtCrtFldLisTbl(A As Database, T, TarTbl, KyFld, Fld, Sep$, Optional IsMem As Boolean, Optional LisFldNm0$)
Dim LisFldNm$: LisFldNm = DftStr(LisFldNm0, Fld & "Lis")
Dim RsFm As DAO.Recordset, LasK, K, RsTo As DAO.Recordset, Lis$
A.Execute FmtQQ("Select ? into [?] from [?] where False", KyFld, TarTbl, T)
A.Execute FmtQQ("Alter Table [?] add column ? ?", TarTbl, LisFldNm, IIf(IsMem, "Memo", "Text(255)"))
Set RsFm = DbqRs(A, FmtQQ("Select ?,? From [?] order by ?,?", KyFld, Fld, T, KyFld, Fld))
If Not RsAny(RsFm) Then Exit Sub
Set RsTo = DbtRs(A, TarTbl)
With RsFm
    LasK = .Fields(0).Value
    Lis = .Fields(1).Value
    .MoveNext
    While Not .EOF
        K = .Fields(0).Value
        If LasK = K Then
            Lis = Lis & Sep & .Fields(1).Value
        Else
            DrInsRs Array(LasK, Lis), RsTo
            LasK = K
            Lis = .Fields(1).Value
        End If
        .MoveNext
    Wend
End With
DrInsRs Array(LasK, Lis), RsTo
End Sub
Sub DrInsRs(Dr, Rs As DAO.Recordset)
With Rs
    .AddNew
    DrSetRs Dr, Rs
    .Update
End With
End Sub
Sub DbtUpdIdFld(A As Database, T, Fld)
Dim D As New Dictionary, J&, Rs As DAO.Recordset, Id$, V
Id = Fld & "Id"
Set Rs = DbqRs(A, FmtQQ("Select ?,?Id from [?]", Fld, Fld, T))
With Rs
    While Not .EOF
        .Edit
        V = .Fields(0).Value
        If D.Exists(V) Then
            .Fields(1).Value = D(V)
        Else
            .Fields(1).Value = J
            D.Add V, J
            J = J + 1
        End If
        .Update
        .MoveNext
    Wend
End With
End Sub
Function DrpTblSql$(T)
DrpTblSql = "Drop Table [" & T & "]"
End Function
Sub DbtDrp(A As Database, T)
If DbHasTbl(A, T) Then A.Execute DrpTblSql(T)
End Sub
Sub DbttDrp(A As Database, TT)
Dim T
For Each T In CvNy(TT)
    DbtDrp A, T
Next
End Sub

Sub BrwMthLocFx()
FxBrw MthLocFx
End Sub

Function CurXlsFxaPj(Fxa) As VBProject
Stop '
End Function

Function VbeFfnPj(A As Vbe, PjFfn) As VBProject
Stop '
End Function

Sub FxaOpn(A)
CurXls.Workbooks.Open A
End Sub
Function FxaMthFullDrs(Fxa, Optional B As WhMdMth) As Drs
Dim Pj As VBProject
Set Pj = CurXlsFxaPj(Fxa)
If IsNothing(Pj) Then
    FxaOpn Fxa
    Set Pj = VbeFfnPj(CurXls.Vbe, Fxa)
    If IsNothing(Pj) Then Stop
End If
Set FxaMthFullDrs = PjMthFullDrs(Pj, B)
End Function
Function CurVbePjFfnPj(PjFfn) As VBProject
Set CurVbePjFfnPj = VbePjFfnPj(CurVbe, PjFfn)
End Function

Function VbePjFfnPj(A As Vbe, Ffn) As VBProject
Dim P As VBProject
For Each P In A.VBProjects
    If PjFfn(P) = Ffn Then Set VbePjFfnPj = P
Next
End Function
Function VbeHasPjFfn(A As Vbe, Ffn) As Boolean
Dim P As VBProject
For Each P In A.VBProjects
    If PjFfn(P) = Ffn Then VbeHasPjFfn = True: Exit Function
Next
End Function
Function CurVbeHasPjFfn(PjFfn) As Boolean
CurVbeHasPjFfn = VbeHasPjFfn(CurVbe, PjFfn)
End Function
Function CurXlsOpnFxa(Fxa) As VBProject
If Not CurVbeHasPjFfn(Fxa) Then
    CurXls.Workbooks.Open Fxa
End If
Dim O As VBProject
Set O = CurVbePjFfnPj(Fxa): If IsNothing(O) Then Stop
Set CurXlsOpnFxa = O
End Function
Sub FxaEns(A)
Stop '
End Sub
Function FxaCrtPj(A) As VBProject
FxaEns A
Set FxaCrtPj = CurXlsOpnFxa(A)
End Function

Function IsFxa(A) As Boolean
IsFxa = LCase(FfnExt(A)) = ".xlam"
End Function
Function IsFb(A) As Boolean
IsFb = LCase(FfnExt(A)) = ".accdb"
End Function
Function PjFfnPjDte(PjFfn) As Date
Select Case True
Case IsFxa(PjFfn): PjFfnPjDte = FileDateTime(PjFfn)
Case IsFb(PjFfn): PjFfnPjDte = FbPjDte(PjFfn)
Case Else: Stop
End Select
End Function
Function DbqVal(A As Database, Q)
DbqVal = RsVal(DbqRs(A, Q))
End Function
Function RsVal(A As DAO.Recordset)
With A
    If .EOF Then Exit Function
    If .BOF Then Exit Function
    RsVal = .Fields(0).Value
End With
End Function

Function PjFfnMthFullCacheDte(PjFfn) As Date
PjFfnMthFullCacheDte = DbqVal(MthDb, FmtQQ("Select PjDte from Mth where PjFfn='?'", PjFfn))
End Function

Sub PjFfnEnsMthFullCache(PjFfn)
Dim D1 As Date
Dim D2 As Date
    D1 = PjFfnPjDte(PjFfn)
    D2 = PjFfnMthFullCacheDte(PjFfn)
Select Case True
Case D1 = 0:  Stop
Case D2 = 0:
Case D1 = D2: Exit Sub
Case D2 > D1: Stop
End Select
DrsRplDbt PjFfnMthFullDrsFmLive(PjFfn), MthDb, "MthCache", FmtQQ("PjFfn='?'", PjFfn)
End Sub

Function PjFfnApp(PjFfn)
Static Y As New Access.Application
Select Case True
Case IsFxa(PjFfn): CurXlsOpnFxa PjFfn: Set PjFfnApp = CurXls
Case IsFb(PjFfn): Y.OpenCurrentDatabase PjFfn: Set PjFfnApp = Y
Case Else: Stop
End Select
End Function
Sub Z_PjFfnMthFullDrsFmLive()
Dim A As Drs, A1$
A1 = CurPjFfnAy()(0)
Set A = PjFfnMthFullDrsFmLive(A1)
WsVis DrsWs(A)
End Sub
Function DrsAddCol(A As Drs, ColNm$, ColVal) As Drs
Dim Fny$(): Fny = A.Fny
Dim NewFny$(): NewFny = Fny: PushI NewFny, ColNm
Set DrsAddCol = Drs(NewFny, DryAddCol(A.Dry, ColVal))
End Function
Function PjFfnMthFullDrsFmLive(PjFfn) As Drs
Dim V As Vbe, A, P As VBProject, PjDte As Date
Set A = PjFfnApp(PjFfn)
Set V = A.Vbe
Set P = VbePjFfnPj(V, PjFfn)
Select Case True
Case IsFb(PjFfn):  PjDte = AcsPjDte(CvAcs(A))
Case IsFxa(PjFfn): PjDte = FileDateTime(PjFfn)
Case Else: Stop
End Select
Set PjFfnMthFullDrsFmLive = DrsAddCol(PjMthFullDrs(P), "PjDte", PjDte)

If IsFb(PjFfn) Then
    CvAcs(A).CloseCurrentDatabase
End If
End Function
Function StruFld(ParamArray Ap()) As Drs
Dim Dry(), Av(), Ele$, LikFldss$, LikFld, X
Av = Ap
For Each X In Av
    LinTRstAsg X, Ele, LikFldss
    For Each LikFld In SslSy(LikFldss)
        PushI Dry, Array(Ele, LikFld)
    Next
Next
Set StruFld = Drs("Ele FldLik", Dry)
End Function

Sub DbStruEns(A As Database, Stru$, B As StruBase)
Dim S$
S = DbtStru(A, LinT1(Stru))
If S = "" Then
    DbStruCrt A, Stru, B
    Exit Sub
End If
If S = Stru Then Exit Sub
DbtReStru A, Stru, B
End Sub

Function AyIntersect(A, B)
If Sz(A) = 0 Or Sz(B) = 0 Then AyIntersect = A: Exit Function
Dim X
For Each X In A
    If AyHas(B, X) Then PushI AyIntersect, X
Next
End Function

Function SelIntoSql$(Sel$, Fm$, Into$)
SelIntoSql = FmtQQ("Select ? from [?] into [?]", Sel, Fm, Into)
End Function

Sub DbtReStru(A As Database, Stru, B As StruBase)
Dim DrpFld$(), NewFld$(), T$, FnyO$(), FnyN$()
T = LinT1(Stru)
FnyO = DbtFny(A, T)
FnyN = StruFny(Stru)

DrpFld = AyMinus(FnyO, FnyN)
NewFld = AyMinus(FnyN, FnyO)
DbtDrpFld A, T, DrpFld
DbtAddFld A, T, NewFld, B
End Sub
Function FdSqlTy$(A As DAO.Field2)
Stop '
End Function
Function FldSqlTy$(F, B As StruBase)
FldSqlTy = FdSqlTy(FldFd(F, "", B))
End Function
Function AddColSql$(T, Fny0, B As StruBase)
Dim O$(), F
For Each F In CvNy(Fny0)
    PushI O, F & " " & FldSqlTy(F, B)
Next
AddColSql = FmtQQ("Alter Table [?] add column ?", T, JnComma(O))
End Function
Sub DbtAddFld(A As Database, T, Fny0, B As StruBase)
If Sz(CvNy(Fny0)) = 0 Then Exit Sub
A.Execute AddColSql(T, Fny0, B)
End Sub

Function DrpFldSql$(T, F)
DrpFldSql = FmtQQ("Alter Table [?] drop column [?]", T, F)
End Function

Sub DbtDrpFld(A As Database, T, Fny0)
Dim F
For Each F In AyNz(CvNy(Fny0))
    A.Execute DrpFldSql(T, F)
Next
End Sub

Sub DbtRen(A As Database, T, TNew$)
FbDb(A.Name).TableDefs(T).Name = TNew
End Sub

Sub EnsMthTbl()
Dim A As Drs
Set A = CurPjFfnAyMthFullDrs
DrsRplDbt A, MthDb, "Mth"
End Sub

Sub DrsRplDbtt(A As Drs, Db As Database, Tny0)
Dim T, Tny$()
Tny = CvNy(Tny0)
For Each T In Tny
    Db.Execute DltSql(T)
Next
DbSqyRun Db, DbDrsNormSqy(Db, A, Tny)
End Sub

Function DbDrsNormSqy(A As Database, B As Drs, Tny$()) As String()

End Function

Function EnsMthFb() As Database
FbEns MthFb
Dim B As StruBase
Set B.F = StruFld("Nm Nm Md Pj", "T50 ToMd", "Txt PjFfn Prm Ret LinRmk", "T3 Ty Mdy", "T4 MdTy", "Int Lno", "Mem NewLines Lines TopRmk")
Const MthCache$ = "MthCache PjFfn Md Nm Ty | Mdy Prm Ret LinRmk TopRmk Lines Lno Pj PjDte MdTy"
Const MthLoc$ = "MthLoc Nm | ToMd NewLines"
Dim Mth$: Mth = "Mth " & LinRmvT1(MthCache)
Dim Db As Database
Set Db = FbDb(MthFb)
DbStruEns Db, MthCache, B
DbStruEns Db, MthLoc, B
DbStruEns Db, Mth, B
Set EnsMthFb = Db
End Function

Function MthDb() As Database
Static A As Database, B As Boolean
If Not B Then
    B = True
    Set A = FbDb(MthFb)
End If
Set MthDb = A
End Function

Function PjFfnMthFullDrs(PjFfn, Optional B As WhMdMth) As Drs
PjFfnEnsMthFullCache PjFfn
Set PjFfnMthFullDrs = PjFfnMthFullDrsFmCache(PjFfn, B)
End Function

Function PjFfnMthFullDrsFmCache(PjFfn, Optional B As WhMdMth) As Drs
Dim Sql$: Sql = FmtQQ("Select * from MthCache where PjFfn='?'", PjFfn)
Set PjFfnMthFullDrsFmCache = DbqDrs(MthDb, Sql)
End Function

Function SampleSq() As Variant()
Const NR% = 10
Const NC% = 10
Dim O(), R%, C%
ReDim O(1 To NR, 1 To NC)
SampleSq = O
For R = 1 To NR
    For C = 1 To NC
        O(R, C) = R * 1000 + C
    Next
Next
SampleSq = O
End Function

Function SampleSqHdr() As Variant()
Const NC% = 10
Dim J%
For J = 0 To NC - 1
    PushI SampleSqHdr, Chr(Asc("A") + J)
Next
End Function

Function SqInsDr(A, Dr, Optional Row& = 1)
Dim O(), C%, R&, NC%, NR&
NC = SqNCol(A)
NR = SqNRow(A)
ReDim O(1 To NR + 1, 1 To NC)
For R = 1 To Row - 1
    For C = 1 To NC
        O(R, C) = A(R, C)
    Next
Next
For C = 1 To NC
    O(Row, C) = Dr(C - 1)
Next
For R = NR To Row Step -1
    For C = 1 To NC
        O(R + 1, C) = A(R, C)
    Next
Next
SqInsDr = O
End Function

Function SampleSqWithHdr() As Variant()
SampleSqWithHdr = SqInsDr(SampleSq, SampleSqHdr)
End Function

Function SampleLo() As ListObject
Set SampleLo = RgLo(SqRg(SampleSqWithHdr, NewA1), "T_Sample")
End Function

Function LoVis(A As ListObject) As ListObject
A.Application.Visible = True
Set LoVis = A
End Function

Function SampleLoVis() As ListObject
Set SampleLoVis = LoVis(SampleLo)
End Function

Function LoDtaAdr$(A As ListObject)
LoDtaAdr = RgAdr(A.DataBodyRange)
End Function
Function RgAdr$(A As Range)
RgAdr = "'" & RgWs(A).Name & "'!" & A.Address
End Function

Function LoPc(A As ListObject) As PivotCache
Dim O As PivotCache
Set O = LoWb(A).PivotCaches.Create(xlDatabase, A.Name, 6)
O.MissingItemsLimit = xlMissingItemsNone
Set LoPc = O
End Function

Function AyNxtNm$(A, Nm$, Optional MaxN% = 99)
If Not AyHas(A, Nm) Then AyNxtNm = Nm: Exit Function
Dim J%, O$
For J = 1 To MaxN
    O = Nm & Format(J, "00")
    If Not AyHas(A, O) Then AyNxtNm = O: Exit Function
Next
Stop
End Function

Function WsPtNy(A As Worksheet) As String()
Dim Pt As PivotTable
For Each Pt In A.PivotTables
    PushI WsPtNy, Pt.Name
Next
End Function

Function WbPtNy(A As Workbook) As String()
Dim Ws As Worksheet
For Each Ws In A.Sheets
    PushIAy WbPtNy, WsPtNy(Ws)
Next
End Function

Function LoPtNm$(A As ListObject)
If Left(A.Name, 2) <> "T_" Then Stop
Dim O$: O = "P_" & Mid(A.Name, 3)
LoPtNm = AyNxtNm(WbPtNy(LoWb(A)), O)
End Function

Sub Z_LoPt()
Dim At As Range, Lo As ListObject
Set Lo = SampleLo
Set At = RgVis(WsA1(WbAddWs(LoWb(Lo))))
PtVis LoPt(Lo, At, "A B", "C D", "F", "E")
Stop
End Sub

Function PtPf(A As PivotTable, F) As PivotField
Set PtPf = A.PivotFields(F)
End Function

Sub PtFldssSetOri(A As PivotTable, Fldss$, Ori As XlPivotFieldOrientation)
Dim F, J%, T
T = Array(False, False, False, False, False, False, False, False, False, False, False, False)
J = 1
For Each F In AyNz(SslSy(Fldss))
    With PtPf(A, F)
        .Orientation = Ori
        .Position = J
        If Ori = xlColumnField Or Ori = xlRowField Then
            .Subtotals = T
        End If
    End With
    J = J + 1
Next
End Sub

Function PtVis(A As PivotTable) As PivotTable
A.Application.Visible = True
Set PtVis = A
End Function

Function PtWb(A As PivotTable) As Workbook
Set PtWb = WsWb(PtWs(A))
End Function

Function LoPt(A As ListObject, At As Range, Rowss$, Dtass$, Optional Colss$, Optional Pagss$) As PivotTable
If LoWb(A).FullName <> RgWb(At).FullName Then Stop: Exit Function
Dim O As PivotTable
Set O = LoPc(A).CreatePivotTable(TableDestination:=At, TableName:=LoPtNm(A))
With O
    .ShowDrillIndicators = False
    .InGridDropZones = False
    .RowAxisLayout xlTabularRow
End With
O.NullString = ""
PtFldssSetOri O, Rowss, xlRowField
PtFldssSetOri O, Colss, xlColumnField
PtFldssSetOri O, Pagss, xlPageField
PtFldssSetOri O, Dtass, xlDataField
Set LoPt = O
End Function

Sub RfhMthLocDb()
If Not IsFfnExist(MthLocFx) Then MsgBox "[MthLocFx] not exist": Exit Sub
DrsInsUpdDbt UsrEdtMthLocDrs, MthDb, "MthLoc"
End Sub

Function MthLocFny() As String()
MthLocFny = SslSy("Nm ToMd")
End Function

Sub UsrEdtMthLocBrw()
Brw DrsFmtss(UsrEdtMthLocDrs)
End Sub

Function IsFfnExist(A) As Boolean
IsFfnExist = Fso.FileExists(A)
End Function
Sub FxCrt(A)
WbSavAs(NewWb, A).Close
End Sub
Function FxEns$(A$)
If Not IsFfnExist(A) Then FxCrt A
FxEns = A
End Function

Sub Z_UsrEdtMthLocDrs()
DrsFmtssBrw UsrEdtMthLocDrs
End Sub
Function UsrEdtMthLocDrs() As Drs
Set UsrEdtMthLocDrs = LoSel(WbLo(FxWb(FxEns(MthLocFx)), "T_UsrEdtMthLoc"), "Nm ToMd")
End Function

Function SqNRow&(A)
On Error Resume Next
SqNRow = UBound(A, 1)
End Function

Function SqNCol&(A)
On Error Resume Next
SqNCol = UBound(A, 1)
End Function

Function SqSel(A, ColIxAy&()) As Variant()
Dim R&
For R = 1 To SqNRow(A)
    Push SqSel, SqRowSel(A, R, ColIxAy)
Next
End Function

Function SqRowSel(A, R&, ColIxAy&()) As Variant()
Dim Ix
For Each Ix In ColIxAy
    Push SqRowSel, A(R, Ix + 1)
Next
End Function

Function LoSq(A As ListObject) As Variant()
If Not IsNothing(A) Then LoSq = A.DataBodyRange.Value
End Function

Function LoSel(A As ListObject, Fldss$) As Drs
Dim Fny$(): Fny = SslSy(Fldss)
Set LoSel = Drs(Fny, SqSel(LoSq(A), AyIxAy(LoFny(A), Fny)))
End Function

Function CrtSkSql$(T, Sk0)
CrtSkSql = FmtQQ("Create unique Index SecondaryKey on [?] (?)", T, JnComma(CvNy(Sk0)))
End Function
Function CrtPkSql$(T)
CrtPkSql = FmtQQ("Create Index PrimaryKey on [?] (?Id) with Primary", T, T)
End Function
Function FxMthMdDrs(A$) As Drs
Dim Wb As Workbook
Set Wb = FxWb(A)
Set FxMthMdDrs = WbMthMdDrs(Wb)
Wb.Close False
End Function

Function WbMthMdDrs(A As Workbook) As Drs
Dim Lo As ListObject
Set Lo = WbLo(A, "T_MthMd")
If IsNothing(Lo) Then Exit Function
If Not IsEqAy(LoFny(Lo), SslSy("Mth Md")) Then Stop
Set WbMthMdDrs = LoDrs(Lo)
End Function
Function CurPjFfnAyMthFullDrs(Optional B As WhPjMth) As Drs
Set CurPjFfnAyMthFullDrs = PjFfnAyMthFullDrs(CurPjFfnAy, B)
End Function
Function PjFfnAyMthFullDrs(PjFfnAy, Optional B As WhPjMth) As Drs
Dim PjFfn
For Each PjFfn In PjFfnAy
    PushDrs PjFfnAyMthFullDrs, PjFfnMthFullDrs(PjFfn, B)
Next
End Function

Function PjFfnAyMthFullWs(PjFfnAy, Optional B As WhPjMth) As Worksheet
Set PjFfnAyMthFullWs = DrsWs(PjFfnAyMthFullDrs(PjFfnAy, B))
End Function
Function PjFfnAyMthFullWb(PjFfnAy$(), Optional B As WhPjMth) As Workbook
Set PjFfnAyMthFullWb = MthFullWbFmt(WsWb(PjFfnAyMthFullWs(PjFfnAy, B)))
End Function

Sub CurPjFfnAyMthFullWb(Optional A As WhPjMth)
WbVis PjFfnAyMthFullWb(CurPjFfnAy, A)
End Sub

Sub Z_CurPjFfnAyMthFullWb()
WbVis PjFfnAyMthFullWb(CurPjFfnAy, WhPjMth(MdMth:=WhMdMth(WhMd("Std"))))
End Sub

Function MthFullWbFmt(A As Workbook) As Workbook
Dim Ws As Worksheet, Lo As ListObject
Set Ws = WbCdNmWs(A, "MthLoc"): If IsNothing(Ws) Then Stop
Set Lo = WsLo(Ws, "T_MthLoc"): If IsNothing(Lo) Then Stop
Dim Ws1 As Worksheet:  GoSub X_Ws1
Dim Pt1 As PivotTable: GoSub X_Pt1
Dim Lo1 As ListObject: GoSub X_Lo1
Dim Pt2 As PivotTable: GoSub X_Pt2
Dim Lo2 As ListObject: GoSub X_Lo2
Ws1.Outline.ShowLevels , 1
Set MthFullWbFmt = WsWb(Ws)
Exit Function
X_Ws1:
    Set Ws1 = WbAddWs(WsWb(Ws))
    Ws1.Outline.SummaryColumn = xlSummaryOnLeft
    Ws1.Outline.SummaryRow = xlSummaryBelow
    Return
X_Pt1:
    Set Pt1 = LoPt(Lo, WsA1(Ws1), "MdTy Nm VbeLinesId Lines", "Pj")
    PtSetRowssOutLin Pt1, "Lines"
    PtSetRowssColWdt Pt1, "VbeLinesId", 12
    PtSetRowssColWdt Pt1, "Nm", 30
    PtSetRowssRepeatLbl Pt1, "MdTy Nm"
    Return
X_Lo1:
    Set Lo1 = PtCpyToLo(Pt1, Ws1.Range("G1"), LoNm:="T_MthLines")
    LoSetColWdt Lo1, "Nm", 30
    LoSetColWdt Lo1, "Lines", 100
    LoSetOutLin Lo1, "Lines"
    
    Return
X_Pt2:
    Set Pt2 = LoPt(Lo1, Ws1.Range("M1"), "MdTy Nm", "Lines")
    PtSetRowssRepeatLbl Pt2, "MdTy"
    Return
X_Lo2:
    Set Lo2 = PtCpyToLo(Pt2, Ws1.Range("Q1"), "T_UsrEdtMthLoc")
    Return
Set MthFullWbFmt = A
End Function

Sub Z_MdMthFullDrs()
DrsBrw MdMthFullDrs(CurMd)
End Sub

Function MdMthFullDry(A As CodeModule, Optional B As WhMth) As Variant()
Dim P As VBProject, Ffn$, Pj$, ShtTy$, Md$, MdTy$
Set P = MdPj(A)
Ffn$ = PjFfn(P)
Pj = P.Name
MdTy = MdShtTy(A)
Md = MdNm(A)
MdMthFullDry = DryInsC4(SrcMthFullDry(MdSrc(A)), Ffn, Pj, MdTy, Md)
End Function

Sub Z_MthFullWbFmt()
Dim Wb As Workbook
Const Fx$ = "C:\Users\user\Desktop\Vba-Lib-1\Mth.xlsx"
MthFullWbFmt WbVis(FxWb(Fx))
Stop
End Sub

Function CurPjFfnAyMthFullWs() As Worksheet
Set CurPjFfnAyMthFullWs = PjFfnAyMthFullWs(CurPjFfnAy)
End Function

Function FbvbeAyMthFullWs(FbvbeAy(), Optional B As WhPjMth) As Worksheet
Dim O As Drs
Set O = FbvbeAyMthFullDrs(FbvbeAy, B)
Set O = DrsAddValIdCol(O, "Nm", "VbeMth")
Set O = DrsAddValIdCol(O, "Lines", "Vbe")
Set FbvbeAyMthFullWs = WsSetCdNmAndLoNm(DrsWs(O), "MthLoc")
End Function



Function MdMthFullDrsFny() As String()
MdMthFullDrsFny = AyAdd(SslSy("PjFfn Pj MdTy Md"), SrcMthIxFullDrFny)
End Function

Function MdMthFullDrs(A As CodeModule, Optional B As WhMth) As Drs
Set MdMthFullDrs = Drs(MdMthFullDrsFny, MdMthFullDry(A, B))
End Function

Function PjMthFullDry(A As VBProject, Optional B As WhMdMth) As Variant()
Dim M
For Each M In AyNz(PjMdAy(A, WhMdMth_Md(B)))
    PushIAy PjMthFullDry, MdMthFullDry(CvMd(M), WhMdMth_Mth(B))
Next
End Function

Function PjMthFullDrs(A As VBProject, Optional B As WhMdMth) As Drs
Dim O As Drs
Set O = Drs(MdMthFullDrsFny, PjMthFullDry(A, B))
Set O = DrsAddValIdCol(O, "Lines", "Pj")
Set O = DrsAddValIdCol(O, "Nm", "PjMth")
Set PjMthFullDrs = O
End Function

Function VbeMthFullWs(A As Vbe, Optional B As WhPjMth) As Worksheet
Set VbeMthFullWs = DrsWs(VbeMthFullDrs(A, B))
End Function

Function VbeMthFullDrs(A As Vbe, Optional B As WhPjMth) As Drs
Dim P
For Each P In AyNz(VbePjAy(A, WhPjMth_Nm(B)))
    PushDrs VbeMthFullDrs, PjMthFullDrs(CvPj(P), WhPjMth_MdMth(B))
Next
End Function

Function FbMthFullDrs(A, Optional B As WhPjMth) As Drs
If False Then
    Set FbMthFullDrs = VbeMthFullDrs(FbAcs(A).Vbe, B)
    Exit Function
End If
Dim Acs As New Access.Application
Debug.Print "FbMthFullDry: "; Now; " Start get Drs "; A; "==============="
Debug.Print "FbMthFullDry: "; Now; " Start open"
Set Acs = FbAcs(A)
Debug.Print "FbMthFullDry: "; Now; " Start get Drs "
Set FbMthFullDrs = VbeMthFullDrs(Acs.Vbe, B)
Debug.Print "FbMthFullDry: "; Now; " Start quit acs "
Acs.Quit acQuitSaveNone
Debug.Print "FbMthFullDry: "; Now; " acs is quit"
Set Acs = Nothing
Debug.Print "FbMthFullDry: "; Now; " acs is nothing"
End Function

Function SrcMthFullDry(A$()) As Variant()
Dim Ix
For Each Ix In AyNz(SrcMthIx(A))
    PushI SrcMthFullDry, SrcMthIxFullDr(A, Ix)
Next
Dim Dr(): GoSub X
If Sz(Dr) > 0 Then
    PushI SrcMthFullDry, Dr
End If
Exit Function
X:
    Dim Dcl$, Cnt%
    Dcl = SrcDclLines(A)
    Cnt = LinCnt(Dcl)
    Const Fldss$ = "Ty Nm Cnt Lines"
    Dim Vy(): Vy = Array("Dcl", "*Dcl", Cnt, Dcl)
    If Dcl = "" Then
        Erase Dr
    Else
        Dr = VyDr(Vy, Fldss, SrcMthIxFullDrFny)
    End If
    Return
End Function
Function FbvbeAyMthFullDrs(FbvbeAy(), Optional B As WhPjMth) As Drs
Dim I, Fst As Boolean
Fst = True
For Each I In FbvbeAy
    Dim A As Drs: GoSub X_A
    If Fst Then
        Set FbvbeAyMthFullDrs = A
        Fst = False
    Else
        PushDrs FbvbeAyMthFullDrs, A
    End If
Next
Exit Function
X_A:
    Select Case True
    Case IsStr(I):            Set A = FbMthFullDrs(I, B)
    Case TypeName(I) = "VBE": Set A = VbeMthFullDrs(CvVbe(I), B)
    Case Else: Stop
    End Select
    Return
End Function


