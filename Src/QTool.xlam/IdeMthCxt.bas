Attribute VB_Name = "IdeMthCxt"
Option Explicit

Function MthCxtFTNoAy(A As Mth) As FTNo()
MthCxtFTNoAy = FTIxNoAy(SrcMthCxtFTIxAy(MdSrc(A.Md), A.Nm))
End Function

Function MthLyCxt(MthLy$()) As String()
MthLyCxt = XX(MthLy, FTNo(1, Sz(MthLy)))
End Function

Private Function SrcMthCxtFTIxAy(A$(), MthNm$) As FTIx()
Dim P() As FTIx
Dim Ay() As FTIx: Ay = SrcMthNmFTIxAy(A, MthNm)
SrcMthCxtFTIxAy = AyMapPXInto(Ay, "XX", A, P)
End Function

Private Function XX(Src$(), X As FTIx) As FTIx
'Src -> X:MthFmno -> MthCxtFTNo
Dim Ix%
For Ix = X.Fmix To X.Toix
    If Not LasChr(Src(Ix)) = "_" Then
        Ix = Ix + 1
        Exit For
    End If
Next
Set XX = FTIx(Ix, X.Toix - 1)
End Function

Private Sub _
ZZ_MthCxtFTNoAy _
 _
()

Dim I
For Each I In MthCxtFTNoAy(CurMth)
    Debug.Print CvFTNo(I).ToStr
Next

End Sub
