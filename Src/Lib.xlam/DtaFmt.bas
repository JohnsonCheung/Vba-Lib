Attribute VB_Name = "DtaFmt"
Option Explicit
Enum eShwZer
    eNo   'Rule:Enum: All Public-Enum-Nm should eXxxXxx format e + CmlStr
    eYes  'Rule:Enum: All Enum-Mbr-Nm should e{CmlShtNm}XXX
            'Rule:Enum: All eXxx{No|Yes} order place No first, so that the default of Enum take No first, because first Enum value is zero
            'Rule:Enum: don't use value
            'Rule:Enum: this is for function option.
            'Rule:Enum: Using these rule will benefit in calling the optional enum as paramter, just giving the enum-mbr will be meaningfull.
End Enum

Function DrFmtCell(Dr, ShwZer As eShwZer) As String()
Dim I, O$()
For Each I In Dr
    Push O, FmtCell(I, ShwZer)
Next
DrFmtCell = O
End Function

Function DrsLy(A As Drs, Optional MaxColWdt& = 100, Optional BrkColNm$) As String()
If AyIsEmp(A.Fny) Then Exit Function
Dim Drs As Drs:
    Drs = DrsAddRowIxCol(A)
Dim BrkColIx%
    BrkColIx = AyIx(A.Fny, BrkColNm)
    If BrkColIx >= 0 Then BrkColIx = BrkColIx + 1 ' Need to increase by 1 due the Ix column is added
Dim Dry(): Dry = Drs.Dry
Push Dry, Drs.Fny
Dim Ay$(): Ay = DryLy(Dry, MaxColWdt, BrkColIx:=BrkColIx)
Dim Lin$: Lin = Pop(Ay)
Dim Hdr$: Hdr = Pop(Ay)
Dim O$()
    PushAy O, Array(Lin, Hdr)
    PushAy O, Ay
    Push O, Lin
'If BrkColNm <> "" Then O = DrsLyInsBrkLin(O, BrkColNm)
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

Function FmtCell$(V, ShwZer As eShwZer)
Dim O$
If ValIsEmp(V) Then Exit Function
If IsObject(V) Then
    FmtCell = "[" & TypeName(V) & "]"
    Exit Function
End If
If ValIsBool(V) Then
    FmtCell = IIf(V, "TRUE", "FALSE")
    Exit Function
End If

If ShwZer = eNo Then
    If IsNumeric(V) Then
        If V = 0 Then Exit Function
    End If
End If
If IsArray(V) Then
    If AyIsEmp(V) Then Exit Function
    FmtCell = "Ay" & UB(V) & ":" & V(0)
    Exit Function
End If
If InStr(V, vbCrLf) > 0 Then
    FmtCell = Brk(V, vbCrLf).S1 & "|.."
    Exit Function
End If
FmtCell = Nz(V, "")
End Function

Private Sub DrsLyInsBrkLin__Tst()
Dim TblLy$()
Dim Act$()
Dim Exp$()
TblLy = FtLy(TstResPth & "DrsLyInsBrkLin.txt")
Act = DrsLyInsBrkLin(TblLy, "Tbl")
Exp = FtLy(TstResPth & "DrsLyInsBrkLin_Exp.txt")
AyPair_EqChk Exp, Act
End Sub

Sub Tst__VbFmt()
DrsLyInsBrkLin__Tst
End Sub
