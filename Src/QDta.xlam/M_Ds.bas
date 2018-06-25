Attribute VB_Name = "M_Ds"
Option Explicit

Function Wb(Optional Vis As Boolean) As Workbook
Dim O As Workbook
Set O = NewWb
With WbFstWs(O)
   .Name = "Ds"
   .Range("A1").Value = B_DsNm
End With
Stop '
'If Not DsIsEmp(A) Then
'   Dim J%
'   For J = 0 To DsNDt(A) - 1
'       WbAddDt O, A.DtAy(J)
'   Next
'End If
If Vis Then WbVis O
Set Wb = O
End Function

Function Ws(Optional Hid As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs(Vis:=Not Hid)
Stop '
'WsA1(O).Value = "*Ds " & A.DsNm
Dim At As Range, J%
Set At = WsRC(O, 2, 1)
Stop '
'For J = 0 To DsNDt(A)
'    Set At = DtAt_NxtAt(A.DtAy(J), At, J)
'Next
Set Ws = O
End Function

Private Sub ZZ_Wb()
Dim Wb As Workbook
Stop
'Set Wb = DsWb(DbDs(CurDb, "Permit PermitD"))
WbVis Wb
Stop
Wb.Close False
End Sub

Sub DsAddDt(ODs As Ds, T As Dt)
Stop '
'If DsHasDt(ODs, T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", ODs.DsNm, T.DtNm)
'Dim N%: N = DtAySz(ODs.DtAy)
'ReDim Preserve ODs.DtAy(N)
'ODs.DtAy(N) = T
End Sub

Sub DsBrw(A As Ds)
AyBrw DsLy(A)
End Sub

Sub DsDmp(A As Ds)
AyDmp DsLy(A)
End Sub

Function DsHasDt(A As Ds, DtNm) As Boolean
If DsIsEmp(A) Then Exit Function
Dim J%
Stop '
'For J = 0 To UBound(A.DtAy)
'    If A.DtAy(J).DtNm = DtNm Then DsHasDt = True: Exit Function
'Next
End Function

Function DsIsEmp(A As Ds) As Boolean
Stop '
'DsIsEmp = DtAy_IsEmp(A.DtAy)
End Function

Function DsLy(A As Ds, Optional MaxColWdt& = 1000, Optional DtBrkLinMapStr$) As String()
Dim O$()
Stop '
'    Push O, "*Ds " & A.DsNm & "=================================================="
Dim Dic As Dictionary ' DicOf_Tn_to_BrkColNm
Stop
'    Set Dic = MapStr_Dic(DtBrkLinMapStr)
'If Not DtAy_IsEmp(A.DtAy) Then
'    Dim J%, DtNm$, Dt As Dt, BrkColNm$
'    For J = 0 To UBound(A.DtAy)
'        Dt = A.DtAy(J)
'        DtNm$ = Dt.DtNm
'        If Dic.Exists(DtNm) Then BrkColNm = Dic(DtNm) Else BrkColNm = ""
'        PushAy O, DtLy(Dt, MaxColWdt, BrkColNm)
'    Next
'End If
'DsLy = O
End Function


