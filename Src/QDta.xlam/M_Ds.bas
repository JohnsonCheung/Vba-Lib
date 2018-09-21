Attribute VB_Name = "M_Ds"
Option Explicit

Function DsAddDt(O As Ds, T As Dt) As Ds
If DsHasDt(O, T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", O.DsNm, T.DtNm)
Dim N%: N = Sz(O.DtAy)
Dim Ay() As Dt
    Ay = O.DtAy
ReDim Preserve Ay(N)
Set Ay(N) = T
Set DsAddDt = Ds(Ay, O.DsNm)
End Function

Function DsAddDtAy(O As Ds, DtAy) As Ds
Dim I
For Each I In DtAy
    Set O = DsAddDt(O, CvDt(I))
Next
End Function

Sub DsBrw(A As Ds)
AyBrw DsLy(A)
End Sub

Sub DsDmp(A As Ds)
AyDmp DsLy(A)
End Sub

Function DsHasDt(A As Ds, DtNm) As Boolean
If DsIsEmp(A) Then Exit Function
Dim Ay() As Dt: Ay = A.DtAy
Dim J%
For J = 0 To UB(Ay)
    If Ay(J).DtNm = DtNm Then DsHasDt = True: Exit Function
Next
End Function

Function DsIsEmp(A As Ds) As Boolean
DsIsEmp = Sz(A.DtAy) = 0
End Function

Function DsLy(A As Ds, Optional MaxColWdt& = 1000, Optional DtBrkLinMapStr$) As String()
Dim O$()
    Push O, "*Ds " & A.DsNm & "=================================================="
Dim Dic As Dictionary ' DicOf_Tn_to_BrkColNm
    Set Dic = MapVbl_Dic(DtBrkLinMapStr)
If Not DsIsEmp(A) Then
    Dim J%, DtNm$, Dt As Dt, BrkColNm$, Ay() As Dt
    Ay = A.DtAy
    For J = 0 To UBound(A.DtAy)
        Set Dt = Ay(J)
        DtNm = Dt.DtNm
        If Dic.Exists(DtNm) Then BrkColNm = Dic(DtNm) Else BrkColNm = ""
        PushAy O, DtLy(Dt, MaxColWdt, BrkColNm)
    Next
End If
DsLy = O
End Function

Function DsWb(A As Ds, Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
With WbFstWs(O)
   .Name = "Ds"
   .Range("A1").Value = A.DsNm
End With
Stop '
'If Not DsIsEmp(A) Then
'   Dim J%
'   For J = 0 To DsNDt(A) - 1
'       WbAddDt O, A.DtAy(J)
'   Next
'End If
If Vis Then WbVis O
Set DsWb = O
End Function

Function DsWs(A As Ds) As Worksheet
Dim O As Worksheet: Set O = NewWs
WsA1(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
Dim Ay() As Dt: Ay = A.DtAy
For J = 0 To DsNDt(A) - 1
    Set At = DtAt(Ay(J), At, J)
Next
Set DsWs = O
End Function
Function DsNDt%(A As Ds)
DsNDt = Sz(A.DtAy)
End Function
Sub ZZ_DsWs()
WsVis DsWs(SampleDs)
End Sub
Sub ZZ_DsLy()
AyDmp DsLy(SampleDs)
End Sub

Private Sub ZZ_DsWb()
Dim Wb As Workbook
Stop '
'Set Wb = DsWb(DbDs(CurDb, "Permit PermitD"))
WbVis Wb
Stop
Wb.Close False
End Sub
