Attribute VB_Name = "F_Ide_MthCxt"
Option Explicit

Function MthCxtFmToLnoAy(A As Mth) As FmToLno()
MthCxtFmToLnoAy = SrcMthCxtFmToLnoAy(MdSrc(A.Md), A.Nm)
End Function

Function MthLy_MthCxtLy(A$()) As String()
MthLy_MthCxtLy = XX(A, FmToLno(1, Sz(A)))
End Function

Private Function SrcMthCxtFmToLnoAy(A$(), MthNm$) As FmToLno()
Dim P() As FmToLno
Dim Ay() As FmToLno: Ay = SrcMthFmToLnoAy(A, MthNm)
SrcMthCxtFmToLnoAy = AyMapPXInto(Ay, "XX", A, P)
End Function

Private Function XX(Src$(), X As FmToLno) As FmToLno
'Src -> X:MthFmLno -> MthCxtFmToLno
Dim FmLno%
For FmLno = X.FmLno To X.ToLno
    If Not LasChr(Src(FmLno - 1)) = "_" Then
        FmLno = FmLno + 1
        Exit For
    End If
Next
Set XX = FmToLno(FmLno, X.ToLno - 1)
End Function

Private Sub _
ZZ_MthCxtFmToLnoAy _
()
Dim I
For Each I In MthCxtFmToLnoAy(CurMth)
    Debug.Print CvFmToLno(I).ToStr
Next
End Sub
