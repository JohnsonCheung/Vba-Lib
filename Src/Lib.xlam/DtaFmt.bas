Attribute VB_Name = "DtaFmt"
Option Explicit
Function DrValCellStr(Dr, ShwZer As Boolean) As String()
Dim I, O$()
For Each I In Dr
    Push O, ValCellStr(I, ShwZer)
Next
DrValCellStr = O
End Function

Function DrsLy(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$, Optional ShwZer As Boolean) As String()
'If BrkColNm changed, insert a break line
If AyIsEmp(A.Fny) Then Exit Function
Dim Drs As Drs:
    Drs = DrsAddRowIxCol(A)
Dim BrkColIx%
    BrkColIx = AyIx(A.Fny, BrkColNm)
    If BrkColIx >= 0 Then BrkColIx = BrkColIx + 1 ' Need to increase by 1 due the Ix column is added
Dim Dry(): Dry = Drs.Dry
Push Dry, Drs.Fny
Dim Ay$(): Ay = DryLy(Dry, MaxColWdt, BrkColIx:=BrkColIx, ShwZer:=ShwZer) '<== Will insert break line if BrkColIx>=0
Dim Lin$: Lin = Pop(Ay)
Dim Hdr$: Hdr = Pop(Ay)
Dim O$()
    PushAy O, Array(Lin, Hdr)
    PushAy O, Ay
    Push O, Lin
DrsLy = O
End Function

Function DrsLyInsBrkLin(TblLy$(), ColNm$) As String()
Dim Hdr$: Hdr = TblLy(1)
Dim Fny$():
    Fny = SplitVBar(Hdr)
    Fny = AyRmvFstEle(Fny)
    Fny = AyRmvLasEle(Fny)
    Fny = SyTrim(Fny)
Dim Ix%
    Ix = AyIx(Fny, ColNm)
Dim DryLy$()
    DryLy = TblLy
    AyWhExclAtCnt DryLy, 0, 2
Dim O$()
    Push O, TblLy(0)
    Push O, TblLy(1)
    PushAy O, DryLy_InsBrkLin(DryLy, Ix)
DrsLyInsBrkLin = O
End Function

Function ValCellStr$(V, ShwZer As Boolean)
'CellStr is a string can be displayed in a cell
Dim O$
If VarIsEmp(V) Then Exit Function
If IsObject(V) Then
    ValCellStr = "[" & TypeName(V) & "]"
    Exit Function
End If
If VarIsBool(V) Then
    ValCellStr = IIf(V, "TRUE", "FALSE")
    Exit Function
End If

If ShwZer Then
    If IsNumeric(V) Then
        If V = 0 Then ValCellStr = "0"
        Exit Function
    End If
End If
If IsArray(V) Then
    If AyIsEmp(V) Then Exit Function
    ValCellStr = "Ay" & UB(V) & ":" & V(0)
    Exit Function
End If
If InStr(V, vbCrLf) > 0 Then
    ValCellStr = Brk(V, vbCrLf).S1 & "|.."
    Exit Function
End If
ValCellStr = Nz(V, "")
End Function

Sub Tst__VbFmt()
DrsLyInsBrkLin__Tst
End Sub

Private Sub DrsLyInsBrkLin__Tst()
Dim TblLy$()
Dim Act$()
Dim Exp$()
TblLy = FtLy(TstResPth & "DrsLyInsBrkLin.txt")
Act = DrsLyInsBrkLin(TblLy, "Tbl")
Exp = FtLy(TstResPth & "DrsLyInsBrkLin_Exp.txt")
AyPair_EqChk Exp, Act
End Sub
