Attribute VB_Name = "IdeMdSrt"
Option Explicit
Type MdSrtRpt
MdIdxDt As Dt
RptDic As Dictionary ' K is Module Name, V is DicCmpRsltLy
End Type
Function SrcSrtCmpLy(A$()) As String()
Dim R As DCRslt: R = SrcSrtCmpRslt(A)
SrcSrtCmpLy = DCRsltLy(R)
End Function

Function SrcSrtCmpRslt(A$()) As DCRslt
Dim A1 As DCRslt
Dim S$(): S = SrcSrtedLy(A)
Dim B As Dictionary
Dim C As Dictionary
Set B = SrcDic(A)
Set C = SrcDic(S)
SrcSrtCmpRslt = DicCmpRslt(B, C)
End Function


Function SrcSrtedBdyLines$(A$())
If AyIsEmp(A) Then Exit Function
Dim Drs As Drs
   Drs = SrcMthDrs(A, WithBdyLines:=True)
Dim MthLinesAy$()
   MthLinesAy = DrsStrCol(Drs, "BdyLines")
Dim I&()
   Dim Ky$(): Ky = MthDrs_SortingKeyAy(Drs)
   I = AySrtInToIxAy(Ky)
Dim O$()
Dim J%
   For J = 0 To UB(I)
       Push O, vbCrLf & MthLinesAy(I(J))
   Next
SrcSrtedBdyLines = JnCrLf(O)
End Function

Private Sub MdSrt__Tst()
Dim Md As CodeModule
GoSub X0
Exit Sub
X0:
    Dim I
    For Each I In PjMdAy(CurPj)
        Set Md = I
        If MdNm(Md) = "Str_" Then
            GoSub Ass
        End If
    Next
    Return
X1:

    Return
Ass:
    Debug.Print MdNm(Md); vbTab;
    Dim BefSrt$(), AftSrt$()
    BefSrt = MdLy(Md)
    AftSrt = SplitCrLf(MdSrtedLines(Md))
    If JnCrLf(BefSrt) = JnCrLf(AftSrt) Then
        Debug.Print "Is Same of before and after sorting ......"
        Return
    End If
    If Not AyIsEmp(AftSrt) Then
        If AyLasEle(AftSrt) = "" Then
            Dim Pfx
            Pfx = Array("There is non-blank-line at end after sorting", "Md=[" & MdNm(Md) & "=====")
            AyBrw AyAddAp(Pfx, AftSrt)
            Stop
        End If
    End If
    Dim A$(), B$(), II
    A = AyMinus(BefSrt, AftSrt)
    B = AyMinus(AftSrt, BefSrt)
    Debug.Print
    If Sz(A) = 0 And Sz(B) = 0 Then Return
    If Not AyIsEmp(AyRmvEmp(A)) Then
        Debug.Print "Sz(A)=" & Sz(A)
        AyBrw A
        Stop
    End If
    If Not AyIsEmp(AyRmvEmp(B)) Then
        Debug.Print "Sz(B)=" & Sz(B)
        AyBrw B
        Stop
    End If
    Return
End Sub
Private Sub MdSrtedLines__Tst()
StrBrw MdSrtedLines(Md("Md_"))
End Sub

Sub PjSrt(A As VBProject)
Dim I
Dim M As CodeModule
For Each I In PjMbrAy(A)
    Set M = I
    MdSrt M
Next
End Sub

Private Sub SrcSrtedBdyLines__Tst()
StrBrw SrcSrtedBdyLines(ZZSrc)
End Sub
Private Function ZZSrc() As String()
ZZSrc = MdSrc(Md("IdeMdSrt"))
End Function

Function SrcSrtedLines$(A$())
Dim O$(), A1$, A2$
A1 = LinesEndTrim(SrcDclLines(A))
A2 = SrcSrtedBdyLines(A)
If LasChr(A2) = vbCr Then Stop
PushNonEmp O, A1
PushNonEmp O, A2
SrcSrtedLines = Join(O, vbCrLf)
End Function



Private Sub SrcSrtedLines__Tst()
StrBrw SrcSrtedLines(ZZSrc)
End Sub

Sub ZZ_SrcSrtCmpLy()
AyBrw SrcSrtCmpLy(ZZSrc)
End Sub

Private Sub ZZ_SrcSrtedBdyLines()
SrcSrtedBdyLines__Tst
End Sub

Private Sub ZZ_SrcSrtedLines()
SrcSrtedLines__Tst
End Sub

Function SrcSrtedLy(A$()) As String()
SrcSrtedLy = SplitCrLf(SrcSrtedLines(A))
End Function

Sub MdSrt(A As CodeModule)
Dim Nm$: Nm = MdNm(A)
Debug.Print "Sorting: "; Nm,
Dim Ay(): Ay = Array("Str_", "Md_", "A_", "SqTp_", "DclOpt")
If AyHas(Ay, Nm) Then
    Debug.Print "<<<< Skipped"
    Exit Sub
End If
Dim Old$: Old = MdLines(A)
Dim NewLines$: NewLines = MdSrtedLines(A)
If Old = NewLines Then
    Debug.Print "<== Same"
    Exit Sub
End If
Debug.Print "<-- Sorted"
MdClr A
A.AddFromString NewLines
Debug.Print "Added...."
MdRmvEndBlankLines A
End Sub

Function MdSrtCmpLy(A As CodeModule) As String()
MdSrtCmpLy = SrcSrtCmpLy(MdSrc(A))
End Function

Function MdSrtCmpRslt(A As CodeModule) As DCRslt
MdSrtCmpRslt = SrcSrtCmpRslt(MdSrc(A))
End Function


Function MdSrtedLines$(A As CodeModule)
MdSrtedLines = SrcSrtedLines(MdSrc(A))
End Function

Function PjMdSrtRpt(A As VBProject) As MdSrtRpt
'SrtCmpDic is a LyDic with Key as MdNm and value is SrtCmpLy
Dim MdAy() As CodeModule: MdAy = PjMdAy(A)
Dim MdNy$(): MdNy = OyNy(MdAy)
Dim LyAy()
Dim IsSam$(), IsDif$(), Sam As Boolean
    Dim J%, R As DCRslt
    For J = 0 To UB(MdAy)
        R = MdSrtCmpRslt(MdAy(J))
        Push LyAy, DCRsltLy(R)
        Sam = DCRsltIsSam(R)
        Push IsSam, IIf(Sam, "*Sam", "")
        Push IsDif, IIf(Sam, "", "*Dif")
    Next
With PjMdSrtRpt
    Set .RptDic = AyPair_Dic(MdNy, LyAy)
    .MdIdxDt = NewDt("Md-Bef-Aft-Srt", "Md Sam Dif", AyZipAp(MdNy, IsSam, IsDif))
End With
End Function
Sub AAA()
ZZ_PjSrtCmpRptWb
End Sub
Sub ZZ_PjSrtCmpRptWb()
Dim O As Workbook: Set O = PjSrtCmpRptWb(CurPj, Vis:=True)
Stop
End Sub
Function PjSrtCmpRptWb(A As VBProject, Optional Vis As Boolean) As Workbook
Dim A1 As MdSrtRpt
A1 = PjMdSrtRpt(A)
Dim O As Workbook: Set O = LyDic_Wb(A1.RptDic)
Dim Ws As Worksheet
Set Ws = WbAddWs(O, "Md Idx", IsBeg:=True)
Dim Lo As ListObject: Set Lo = DtLo(A1.MdIdxDt, WsA1(Ws))
LoCol_LnkWs Lo, "Md"
If Vis Then WbVis O
Set PjSrtCmpRptWb = O
End Function
