Attribute VB_Name = "M_Lin"
Option Explicit

Sub LinAsgTRst(Lin, OTerm, ORst)
With Brk1(Lin, " ")
    OTerm = .S1
    ORst = .S2
End With
End Sub

Sub LinAsgTTRst(Lin, OTerm1, OTerm2, ORst)
Dim A$: A = Lin
OTerm1 = LinShiftTerm(A)
OTerm2 = LinShiftTerm(A)
ORst = A
End Sub

Function LinHasDDRmk(Lin) As Boolean
LinHasDDRmk = HasSubStr(Lin, "--")
End Function

Function LinIsSngTerm(Lin) As Boolean
With Brk1(Lin, " ")
    LinIsSngTerm = .S1 <> "" And .S2 = ""
End With
End Function


Function LinNm$(Lin)
Dim J%
If IsLetter(FstChr(Lin)) Then
   For J = 2 To Len(Lin)
       If Not IsNmChr(Mid(Lin, J, 1)) Then Exit For
   Next
   LinNm = Left(Lin, J - 1)
End If
End Function

Function LinPfxErMsg$(Lin, Pfx$)
If HasPfx(Lin, Pfx) Then Exit Function
LinPfxErMsg = FmtQQ("First Char must be [?]", Pfx)
End Function

Function LinRmvDDRmk$(A)
Dim S$
If LinHasDDRmk(A) Then
    S = ""
Else
    S = A
End If
End Function

Function LinRmvT1$(Lin)
LinRmvT1 = Brk1(Trim(Lin), " ").S2
End Function

Function LinShiftTerm$(OLin)
With Brk1(OLin, " ")
    LinShiftTerm = .S1
    OLin = .S2
End With
End Function

Function LinT1$(Lin)
LinT1 = Brk1(Lin, " ").S1
End Function

Function LinT1Rst(Lin) As T1Rst
Dim O As T1Rst
With Brk1(Lin, " ")
    O.T1 = .S1
    O.Rst = .S2
End With
End Function

Function LinT2$(Lin)
LinT2 = Brk1(Lin, " ").S2
End Function

Sub ZZZ__Tst()
ZZ_LinRmvT1
End Sub

Private Sub ZZ_LinRmvT1()
Ass LinRmvT1("  df dfdf  ") = "dfdf"
End Sub
