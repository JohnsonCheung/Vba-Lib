Attribute VB_Name = "F_Ide_MthRmk"
Option Explicit

Sub MthRmk(A As Mth)
Dim P() As FmToLno: P = MthCxtFmToLnoAy(A)
Dim J%
For J = UB(P) To 0 Step -1
    MthFmToLno_Rmk A, P(J)
Next
End Sub

Sub MthUnRmk(A As Mth)
Dim P() As FmToLno: P = MthCxtFmToLnoAy(A)
Dim J%
For J = UB(P) To 0 Step -1
    MthFmToLno_UnRmk A, P(J)
Next
End Sub

Private Function MthCxtLy_IsRemarked(A$()) As Boolean
If Sz(A) = 0 Then Exit Function
If Not IsPfx(A(0), "Stop '") Then Exit Function
Dim J%
For J = 1 To UB(A)
    If Left(A(J), 1) <> "'" Then Exit Function
Next
MthCxtLy_IsRemarked = True
End Function

Private Sub MthFmToLno_Rmk(A As Mth, X As FmToLno)
Dim Ly$():  Ly = Md_Ly_ByFmToLno(A.Md, X)
If MthCxtLy_IsRemarked(Ly) Then Exit Sub
Dim J%, L$
For J = X.FmLno To X.ToLno
    L = A.Md.Lines(J, 1)
    A.Md.ReplaceLine J, "'" & L
Next
A.Md.InsertLines X.FmLno, "Stop" & " '"
End Sub

Private Sub MthFmToLno_UnRmk(A As Mth, X As FmToLno)
Dim Ly$():  Ly = Md_Ly_ByFmToLno(A.Md, X)
If Not MthCxtLy_IsRemarked(Ly) Then Exit Sub
Dim J%, L$
If Not IsPfx(A.Md.Lines(X.FmLno, 1), "Stop '") Then Stop
For J = X.FmLno + 1 To X.ToLno
    L = A.Md.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.Md.ReplaceLine J, Mid(L, 2)
Next
A.Md.DeleteLines X.FmLno, 1
End Sub

Private Sub ZZ_MthRmk()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "ZZA")
            Debug.Assert LinesVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let ZZA(V)||End Property"
MthRmk M:   Debug.Assert LinesVbl(MthLines(M)) = "Property Get ZZA()|Stop '|End Property||Property Let ZZA(V)|Stop '|'|End Property"
MthUnRmk M: Debug.Assert LinesVbl(MthLines(M)) = "Property Get ZZA()|End Property||Property Let ZZA(V)||End Property"
End Sub
