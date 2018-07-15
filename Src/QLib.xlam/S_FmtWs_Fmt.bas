Attribute VB_Name = "S_FmtWs_Fmt"
Option Explicit
'Private Sub Y(Lo As ListObject, C$(), A As XlTotalsCalculation)
'Dim J%
'For J = 0 To UB(C)
'    Lo.ListColumns(C(J)).TotalsCalculation = A
'Next
'End Sub
'
'Private Function ZTitDry(Fny$()) As Variant()
''From B_Tit & Fny, return TitDry
''If some column has no title, use FldNm as Tit
'Dim O(), J%, Ix%, Ay() As CnoVal
'For J = 0 To UB(Fny)
'    Ix = B_Tit.CnoIx(J)
'    If Ix = -1 Then
'        Push O, Array(Fny(J))
'    Else
'        Push O, SplitVBar(Ay(J).V, Trim:=True)  ' V contains Tit of current fld
'    End If
'Next
'ZTitDry = O
'End Function
'
'Private Function ZTitNRow%()
'Dim O%
'    Dim A$(), J%, M%
'    A$() = B_Tit.StrValAy
'    For J = 0 To UB(A)
'        M = Sz(Split(A(J), "|"))
'        If M > 0 Then O = M
'    Next
'ZTitNRow = O
'End Function
'
'Private Function ZTitSq(Fny$()) As Variant()
'Dim TitSq(): TitSq = DrySq(ZTitDry(Fny))
'ZTitSq = SqTranspose(TitSq)
'End Function
'
'Private Sub ZZ_DoFmt()
'Dim A As New LoFmtr
'Dim B As LoFmtrRslt
'Set B = A.InitBySampleLy.Validate
'B.FmtWs.DoFmt SampleLo
''A.InitBySampleLy.Validate.FmtWs.DoFmt SampleLo
'End Sub
'
