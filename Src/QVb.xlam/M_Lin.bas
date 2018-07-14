Attribute VB_Name = "M_Lin"
Option Explicit

Property Get LinHasDDRmk(Lin) As Boolean
LinHasDDRmk = HasSubStr(Lin, "--")
End Property

Property Get LinIsSngTerm(Lin) As Boolean
With Brk1(Lin, " ")
    LinIsSngTerm = .S1 <> "" And .S2 = ""
End With
End Property

Property Get LinNm$(Lin)
Dim J%
If IsLetter(FstChr(Lin)) Then
   For J = 2 To Len(Lin)
       If Not IsNmChr(Mid(Lin, J, 1)) Then Exit For
   Next
   LinNm = Left(Lin, J - 1)
End If
End Property

Property Get LinPfxErMsg$(Lin, Pfx$)
If HasPfx(Lin, Pfx) Then Exit Property
LinPfxErMsg = FmtQQ("First Char must be [?]", Pfx)
End Property

Property Get LinRmvDDRmk$(A)
Dim S$
If LinHasDDRmk(A) Then
    S = ""
Else
    S = A
End If
Stop '
LinRmvDDRmk = S
End Property

Property Get LinRmvT1$(Lin)
LinRmvT1 = Brk1(Trim(Lin), " ").S2
End Property

Property Get LinShiftTerm$(OLin)
With Brk1(OLin, " ")
    LinShiftTerm = .S1
    OLin = .S2
End With
End Property

Property Get LinT1$(Lin)
LinT1 = Brk1(Lin, " ").S1
End Property

Property Get LinT1Rst(Lin) As T1Rst
Dim O As T1Rst
With Brk1(Lin, " ")
    O.T1 = .S1
    O.Rst = .S2
End With
LinT1Rst = O
End Property

Property Get LinT2$(Lin)
LinT2 = Brk1(Lin, " ").S2
End Property

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

Sub ZZ__Tst()
ZZ_LinRmvT1
End Sub

Private Sub ZZ_LinRmvT1()
Ass LinRmvT1("  df dfdf  ") = "dfdf"
End Sub
