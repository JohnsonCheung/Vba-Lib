Attribute VB_Name = "M_Dt"
Option Explicit

Sub DtBrw(A As Dt, Optional Fnn$)
AyBrw DtLy(A), Dft(Fnn, A.DtNm)
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

Property Get DtSel(A As Dt, ColNy0) As Dt
Dim ReOrdFny$(): ReOrdFny = DftNy(ColNy0)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
Set DtSel = Dt(A.DtNm, OFny, ODry)
End Property

Property Get DtSrt(A As Dt, ColNm$, Optional IsDes As Boolean) As Dt
Set DtSrt = Dt(A.DtNm, A.Fny, DrsSrt(DtDrs(A), ColNm, IsDes).Dry)
End Property

Property Get DtDrpCol(A As Dt, ColNy0, Optional DtNm$) As Dt
Dim A1 As Drs: Set A1 = DrsDrpCol(DtDrs(A), ColNy0)
Set DtDrpCol = Dt(Dft(DtNm, A.DtNm), A1.Fny, A1.Dry)
End Property
Property Get DtLy(A As Dt, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
Dim O$()
   Push O, "*Tbl " & A.DtNm
   PushAy O, DrsLy(DtDrs(A), MaxColWdt, BrkColNm, ShwZer)
DtLy = O
End Property
Function DtAt_NxtAt(A As Dt, At As Range, J%) As Range
At.Value = "(" & J & ") " & A.DtNm
Set At = RgRC(At, 2, 1)
Dim Ly$(): Ly = DrsLy(DtDrs(A))
AyRgV Ly, At
Set At = RgRC(At, 1 + Sz(Ly), 1)
Set DtAt_NxtAt = At
End Function
Function DtPutWb(A As Dt, Wb As Workbook) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(Wb, A.DtNm)
DrsLo DtDrs(A), WsA1(O), A.DtNm
Set DtPutWb = O
End Function

Property Get DtLo(A As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set DtLo = DrsLo(DtDrs(A), R, A.DtNm)
RgRC(R, 0, 1).Value = A.DtNm
End Property

Property Get DtWs(A As Dt, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
DrsLo DtDrs(A), WsA1(O)
Set DtWs = O
If Vis Then WsVis O
End Property

