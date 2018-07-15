Attribute VB_Name = "M_Ds"
Option Explicit

Property Get DsHasDt(A As Ds, DtNm) As Boolean
If DsIsEmp(A) Then Exit Property
Dim Ay() As Dt: Ay = A.DtAy
Dim J%
For J = 0 To UB(Ay)
    If Ay(J).DtNm = DtNm Then DsHasDt = True: Exit Property
Next
End Property

Property Get DsAddDt(O As Ds, T As Dt) As Ds
If DsHasDt(O, T.DtNm) Then Err.Raise 1, , FmtQQ("DsAddDt: Ds[?] already has Dt[?]", O.DsNm, T.DtNm)
Dim N%: N = Sz(O.DtAy)
Dim Ay() As Dt
    Ay = O.DtAy
ReDim Preserve Ay(N)
Set Ay(N) = T
Set DsAddDt = Ds(Ay, O.DsNm)
End Property

Property Get DsIsEmp(A As Ds) As Boolean
DsIsEmp = Sz(A.DtAy) = 0
End Property

Property Get DsLy(A As Ds, Optional MaxColWdt& = 1000, Optional DtBrkLinMapStr$) As String()
Dim O$()
    Push O, "*Ds " & A.DsNm & "=================================================="
Dim Dic As Dictionary ' DicOf_Tn_to_BrkColNm
    Stop '
'    Set Dic = MapStr_Dic(DtBrkLinMapStr)
If Not DsIsEmp(A) Then
    Dim J%, DtNm$, Dt As Dt, BrkColNm$, Ay() As Dt
    Ay = A.DtAy
    For J = 0 To UBound(A.DtAy)
        Set Dt = Ay(J)
        DtNm = Dt.DtNm
        If Dic.Exists(DtNm) Then BrkColNm = Dic(DtNm) Else BrkColNm = ""
        Stop '
        'PushAy O, DtLy(Dt, MaxColWdt, BrkColNm)
    Next
End If
DsLy = O
End Property

Property Get DsWb(A As Ds, Optional Vis As Boolean) As Workbook
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
End Property


Sub DsBrw(A As Ds)
AyBrw DsLy(A)
End Sub

Sub DsDmp(A As Ds)
AyDmp DsLy(A)
End Sub

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
