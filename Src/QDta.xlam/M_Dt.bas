Attribute VB_Name = "M_Dt"
Option Explicit


Sub DtBrw(A As Dt, Optional Fnn$)
AyBrw DtLy(A), Dft(Fnn, DtNm)
End Sub

Property Get DtCsvLy(A As Dt) As String()
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(AyQuoteDbl(A.Fny))
For Each Dr In A.Dry
   Push O, FmtQQAv(QQStr, Dr)
Next
End Property

Sub DtDmp(A As Dt)
AyDmp DtLy(A)
End Sub

Property Get DtDrpCol(A As Dt, ColNy0, Optional DtNm$) As Dt
Dim A As Drs: Set A = DrsDrpCol(ColNy0)
DrpCol = Dt(Dft(DtNm, Me.DtNm), A.Fny, A.Dry)
End Property
Property Get DtLy(A As Dt, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
Dim O$()
   Push O, "*Tbl " & DtNm
   PushAy O, Drs.Ly(MaxColWdt, BrkColNm, ShwZer)
Ly = O
End Property
Function DtAt_NxtAt(A As Dt, At As Range, J%) As Range
At.Value = "(" & J & ") " & A.DtNm
Set At = RgRC(At, 2, 1)
Dim Ly$(): Ly = A.Drs.Ly
AyRgV Ly, At
Set At = RgRC(At, 1 + Sz(Ly), 1)
Set DtAt_NxtAt = At
End Function
Function PutWb(A As Workbook) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(A, DtNm)
Drs.Lo WsA1(O), DtNm
Set WbAddDt = O
End Function

Property Get Lo(At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set Lo = Drs.Lo(R, A.DtNm)
RgRC(R, 0, 1).Value = A.DtNm
End Property

Property Get Ws(Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
Drs.Lo WsA1(O)
Set Ws = O
If Vis Then WsVis O
End Property

Function DtNmDrs_Dt(A$, Drs As Drs) As Dt
With DtNmDrs_Dt
    .DtNm = A
    .Fny = Drs.Fny
    .Dry = Drs.Dry
End With
End Function
Property Get ReOrd(ColNy0) As Dt
Dim ReOrdFny$(): ReOrdFny = DftNy(ColNy0)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = Drx.ReOrd(IxAy)
Dim O As New Dt
Set ReOrd = O.Init(DtNm, OFny, ODry)
End Property

Property Get Srt(ColNm$, Optional IsDes As Boolean) As Dt
Set Srt = DtByDrs(DtNm, Drs.Srt(ColNm, IsDes))
End Property

