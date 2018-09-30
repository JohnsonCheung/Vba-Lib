Attribute VB_Name = "IdeMthLoc"
Option Explicit
Public Const MthLocFb$ = "C:\Users\User\Desktop\Vba-Lib-1\MthLoc.accdb"
Public Const MthLocFx$ = "C:\Users\User\Desktop\Vba-Lib-1\MthLoc.xlsx"

Sub GenAndBrwMthLoc()
WbVis RefreshedUsrEditMthLocWb
End Sub

Sub BrwMthLocFx()
FxBrw MthLocFx
End Sub

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

Sub PtFldssSetOrientation(A As PivotTable, Fldss$, Ori As XlPivotFieldOrientation)
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
PtFldssSetOrientation O, Rowss, xlRowField
PtFldssSetOrientation O, Colss, xlColumnField
PtFldssSetOrientation O, Pagss, xlPageField
PtFldssSetOrientation O, Dtass, xlDataField
Set LoPt = O
End Function

Function RefreshedUsrEditMthLocWb() As Workbook
FfnDlt MthLocFx
Dim Wb As Workbook
'Set Wb = PtWb(LoPt(CurFbvbeAyMthFullWs.ListObjects(1), At, "MdTy Nm NmCnt Mdy", "Pj"))
Set RefreshedUsrEditMthLocWb = WbSavAs(Wb, MthLocFx)
End Function
Sub RfhMthLocDb()
If Not IsFfnExist(MthLocFx) Then MsgBox "[MthLocFx] not exist": Exit Sub
DrsInsUpdDbt UsrEdtMthLocDrs, MthLocDb, "MthLoc"
End Sub
Function MthLocFny() As String()
MthLocFny = SslSy("Nm ToMd")
End Function

Sub UsrEdtMthLocBrw()
Brw DrsFmtss(UsrEdtMthLocDrs)
End Sub
Function IsFfnExist(A$) As Boolean
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

Function MthLocDb() As Database
EnsMthLocFb
Set MthLocDb = FbDb(MthLocFb)
End Function

Sub EnsMthLocFb()
Dim Fb$: Fb = MthLocFb
FbEns Fb
If FbHasTbl(Fb, "MthLoc") Then Exit Sub
With FbDb(Fb)
    .Execute "Create Table MthLoc (Nm Text(100) Not Null,ToMd Text(100))"
    .Execute CrtSkSql("MthLoc", "Nm")
End With
End Sub

Function CrtSkSql$(T, Sk0)
CrtSkSql = FmtQQ("Create unique Index SecondaryKey on [?] (?)", T, JnComma(CvNy(Sk0)))
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
