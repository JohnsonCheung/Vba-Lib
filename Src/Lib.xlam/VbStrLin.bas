Attribute VB_Name = "VbStrLin"
Option Explicit
Type FstTermAyRstAy
    FstTermAy() As String
    RstAy() As String
End Type
Type FstTermRst
    FstTerm As String
    Rst As String
End Type
Function LinIsEmp(A) As Boolean
LinIsEmp = Trim(A) = ""
End Function
Function LyLnxAy(A$()) As Lnx()
Dim J&, O() As Lnx
If AyIsEmp(A) Then Exit Function
For J = 0 To UB(A)
    LnxPush O, NewLnx(J, A(J))
Next
LyLnxAy = O
End Function
Function LinHasDDRmk(I) As Boolean
Dim L$: L = Trim(I)
LinHasDDRmk = True
If L = "" Then Exit Function
If HasPfx(L, "--") Then Exit Function
LinHasDDRmk = False
End Function
Function LyHasMajPfx(A$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(A)
    If HasPfx(A(J), MajPfx) Then Cnt = Cnt + 1
Next
LyHasMajPfx = Cnt > (Sz(A) \ 2)
End Function

Function LnxAy_Ly(A() As Lnx) As String()
Dim J%, O$()
For J = 0 To LnxUB(A)
    Push O, A(J).Lin
Next
LnxAy_Ly = O
End Function

Sub LnxPush(O() As Lnx, M As Lnx)
Dim N&
    N = LnxSz(O)
ReDim Preserve O(N)
    O(N) = M
End Sub
Sub LinAsgTRst(A, OT, ORst)
Dim L$: L = A
OT = LinShiftTerm(L)
ORst = L
End Sub

Function LinBrkFstTermRst(A$) As FstTermRst
Dim O As FstTermRst
With Brk1(A, " ")
    O.FstTerm = .S1
    O.Rst = .S2
End With
LinBrkFstTermRst = O
End Function
Private Sub LyBrkFstTermAyRstAy__Tst()
Dim A$()
Push A, "lskdfj sldkfj sldfj sldkfj sldf j"
Push A, "lksj flskdj flsdjk fsldjkf"
Dim Act As FstTermAyRstAy
Act = LyBrkFstTermAyRstAy(A)
Stop
End Sub
Function LyBrkFstTermAyRstAy(A$()) As FstTermAyRstAy
Dim J&, FstTermAy$(), RstAy$()
For J = 0 To UB(A)
    With LinBrkFstTermRst(A(J))
        Push FstTermAy, .FstTerm
        Push RstAy, .Rst
    End With
Next
With LyBrkFstTermAyRstAy
    .FstTermAy = FstTermAy
    .RstAy = RstAy
End With
End Function
Function LinPfxErMsg$(Lin$, Pfx$)
If HasPfx(Lin, Pfx) Then Exit Function
LinPfxErMsg = FmtQQ("First Char must be [?]", Pfx)
End Function
Sub LinAsgTTRst(A, OT1, OT2, ORst)
Dim L$: L = A
OT1 = LinShiftTerm(L)
OT2 = LinShiftTerm(L)
ORst = L
End Sub
Function LinShiftTerm$(OLin)
With Brk1(OLin, " ")
    LinShiftTerm = .S1
    OLin = .S2
End With
End Function

Function LnxSz%(A() As Lnx)
On Error Resume Next
LnxSz = UBound(A) + 1
End Function

Function LnxUB%(A() As Lnx)
LnxUB = LnxSz(A) - 1
End Function

Function NewLnx(Lx&, Lin$) As Lnx
With NewLnx
    .Lx = Lx
    .Lin = Lin
End With
End Function


Function LinesToVbl$(A$)
If InStr(A, "|") Then Er "LinesToVbl", "Cannt have [|] in {Lines}", A
LinesToVbl = Replace(A, vbCrLf, "|")
End Function


Function LinesAy_Wdt%(A$())
Dim O%, J&, M%
For J = 0 To UB(A)
   M = LinesWdt(A(J))
   If M > O Then O = M
Next
LinesAy_Wdt = O
End Function

Function LinesWdt%(A$)
LinesWdt = AyWdt(SplitCrLf(A))
End Function

Function LinesLasLin$(A$)
LinesLasLin = AyLasEle(SplitCrLf(A))
End Function

Function LinesLinCnt&(A$)
LinesLinCnt = Sz(SplitCrLf(A))
End Function


Function LinNm$(A)
Dim J%
If IsLetter(FstChr(A)) Then
   For J = 2 To Len(A)
       If Not IsNmChr(Mid(A, J, 1)) Then Exit For
   Next
   LinNm = Left(A, J - 1)
End If
End Function

Function LinRmvT1$(A)
LinRmvT1 = Brk1(Trim(A), " ").S2
End Function

Function LinRstTerm$(A)
LinRstTerm = Brk1(Trim(A), " ").S2
End Function

Function LinT1$(A)
LinT1 = Brk1(Trim(A), " ").S1
End Function

Function LinT2$(A)
Dim L$: L = LinRmvT1(A)
LinT2 = LinT1(L)
End Function

Private Sub LinRmvT1__Tst()
Ass LinRmvT1("  df dfdf  ") = "dfdf"
End Sub
