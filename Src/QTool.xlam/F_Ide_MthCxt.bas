Attribute VB_Name = "F_Ide_MthCxt"
Option Explicit

Function MthCxtFTNoAy(A As Mth) As FTNo()
MthCxtFTNoAy = SrcMthCxtFTNoAy(MdSrc(A.Md), A.Nm)
End Function

Function MthLy_MthCXTLy(A$()) As String()
MthLy_MthCXTLy = XX(A, FTNo(1, Sz(A)))
End Function

Private Function SrcMthCxtFTNoAy(A$(), MthNm$) As FTNo()
Dim P() As FTNo
Dim Ay() As FTNo: Ay = SrcMthFTNoAy(A, MthNm)
SrcMthCxtFTNoAy = AyMapPXInto(Ay, "XX", A, P)
End Function

Private Function XX(Src$(), X As FTNo) As FTNo
'Src -> X:MthFmno -> MthCxtFTNo
Dim Fmno%
For Fmno = X.Fmno To X.Tono
    If Not LasChr(Src(Fmno - 1)) = "_" Then
        Fmno = Fmno + 1
        Exit For
    End If
Next
Set XX = FTNo(Fmno, X.Tono - 1)
End Function

Private Sub _
ZZ_MthCxtFTNoAy _
()
Dim I
For Each I In MthCxtFTNoAy(CurMth)
    Debug.Print CvFTNo(I).ToStr
Next
End Sub
