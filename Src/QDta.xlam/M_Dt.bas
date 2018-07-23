Attribute VB_Name = "M_Dt"
Option Explicit

Function DtAt(A As Dt, At As Range, J%) As Range
At.Value = "(" & J & ") " & A.DtNm
Set At = RgRC(At, 2, 1)
Dim Ly$(): Ly = DrsLy(DtDrs(A))
AyRgV Ly, At
Set At = RgRC(At, 1 + Sz(Ly), 1)
Set DtAt = At
End Function

Function DtCsvLy(A As Dt) As String()
Dim O$()
Dim QQStr$
Dim Dr
Push O, JnComma(AyQuoteDbl(A.Fny))
For Each Dr In A.Dry
   Push O, FmtQQAv(QQStr, Dr)
Next
End Function

Function DtDrpCol(A As Dt, ColNy0, Optional DtNm$) As Dt
Dim A1 As Drs: Set A1 = DrsDrpCol(DtDrs(A), ColNy0)
Set DtDrpCol = Dt(Dft(DtNm, A.DtNm), A1.Fny, A1.Dry)
End Function

Function DtDrs(A As Dt) As Drs
Set DtDrs = Drs(A.Fny, A.Dry)
End Function

Function DtIsEmp(A As Dt) As Boolean
DtIsEmp = AyIsEmp(A.Dry)
End Function

Function DtLo(A As Dt, At As Range) As ListObject
Dim R As Range
If At.Row = 1 Then
    Set R = RgRC(At, 2, 1)
Else
    Set R = At
End If
Set DtLo = DrsLo(DtDrs(A), R, A.DtNm)
RgRC(R, 0, 1).Value = A.DtNm
End Function

Function DtLy(A As Dt, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
Dim O$()
   Push O, "*Tbl " & A.DtNm
   PushAy O, DrsLy(DtDrs(A), MaxColWdt, BrkColNm, ShwZer)
DtLy = O
End Function

Function DtPutWb(A As Dt, Wb As Workbook) As Worksheet
Dim O As Worksheet
Set O = WbAddWs(Wb, A.DtNm)
DrsLo DtDrs(A), WsA1(O), A.DtNm
Set DtPutWb = O
End Function

Function DtSel(A As Dt, ColNy0) As Dt
Dim ReOrdFny$(): ReOrdFny = DftNy(ColNy0)
Dim IxAy&(): IxAy = AyIxAy(A.Fny, ReOrdFny)
Dim OFny$(): OFny = AyReOrd(A.Fny, IxAy)
Dim ODry(): ODry = DryReOrd(A.Dry, IxAy)
Set DtSel = Dt(A.DtNm, OFny, ODry)
End Function

Function DtSrt(A As Dt, ColNm$, Optional IsDes As Boolean) As Dt
Set DtSrt = Dt(A.DtNm, A.Fny, DrsSrt(DtDrs(A), ColNm, IsDes).Dry)
End Function

Function DtWs(A As Dt, Optional Vis As Boolean) As Worksheet
Dim O As Worksheet
Set O = NewWs(A.DtNm)
DrsLo DtDrs(A), WsA1(O)
Set DtWs = O
If Vis Then WsVis O
End Function

Sub DtBrw(A As Dt, Optional Fnn$)
AyBrw DtLy(A), Dft(Fnn, A.DtNm)
End Sub

Sub DtDmp(A As Dt)
AyDmp DtLy(A)
End Sub
