Attribute VB_Name = "F_EnsMthMdy"
Option Explicit

Sub MthEnsMdy(A As CodeModule, MthNm$, Mdy$)
Dim Ix%(): Ix = MthLnoAy(A, MthNm)
Dim J%
For J = 0 To UB(Ix)
   MdMthLno_EnsMdy A, Ix(J), Mdy
Next
End Sub

Sub MthEnsPrivate(A As CodeModule, MthNm$)
MthEnsMdy A, MthNm, "Private"
End Sub

Sub MthEnsPublic(A As CodeModule, MthNm$)
MthEnsMdy A, MthNm, "Public"
End Sub

Sub MthEnsMdy__Tst()
Dim M As CodeModule
Dim MthNm$
Dim Mdy$
'Set M = Md("IDE_Feature001_EnsMthMdy")
MthNm = "AAAX"
Mdy = "Public"
MthEnsMdy M, MthNm, Mdy
End Sub

Private Sub MdMthLno_EnsMdy(A As CodeModule, MthLno%, Mdy$)
Dim Lin$
   Lin = A.Lines(MthLno, 1)
If Not SrcLin_IsMth(Lin) Then
   Er "MdMthLno", "Given {Lin} of {Md} of {MthLno} is not a method", Lin, MdNm(A), MthLno
End If
Dim NewLin$
   Select Case Mdy
   Case "Public", "": NewLin = SrcLin_RmvMdy(Lin)
   Case "Private": NewLin = "Private " & SrcLin_RmvMdy(Lin)
   Case Else
       Er "MdMthLno", "Given parament {Mdy} must be ["" | Public | Private]", Mdy
   End Select
If Lin = NewLin Then
   Debug.Print FmtQQ("MdMthLno_EnsMdy: Same Mdy[?] in Lin[?]", Mdy, Lin)
   Exit Sub
End If
MdRplLin A, MthLno, NewLin
Debug.Print FmtQQ("MdMthLno_EnsMdy: Mdy[?] Of MthLno[?] of Md[?] is ensured", Mdy, MthLno, MdNm(A))
Debug.Print FmtQQ("                 OldLin[?]", Lin)
Debug.Print FmtQQ("                 NewLin[?]", NewLin)
End Sub
