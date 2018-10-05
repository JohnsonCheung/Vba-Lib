Attribute VB_Name = "IdeMthCxt"
Option Explicit
Function SrcMthCxtFT(A$(), MthNm$) As FTNo()
Dim P() As FTIx
Dim Ix() As FTIx: Ix = SrcMthNmFT(A, MthNm)
SrcMthCxtFT = AyMapPXInto(Ix, "CxtIx", A, P)
End Function


Function SrcMthFT_CxtFT(Src$(), Mth As FTIx) As FTIx
'Src -> X:MthFmno -> MthCxtFTNo
With Mth
    Dim Ix%
    For Ix = .Fmix To .Toix
        If Not LasChr(Src(Ix)) = "_" Then
            Ix = Ix + 1
            Exit For
        End If
    Next
    Set SrcMthFT_CxtFT = FTIx(Ix, .Toix - 1)
End With
End Function

Private Sub ZZ_MthCxtFT _
 _
()

Dim I
For Each I In MthCxtFT(CurMth)
    Debug.Print CvFTNo(I).ToStr
Next

End Sub
Function MthCxtFT(A As Mth) As FTNo()
MthCxtFT = SrcMthCxtFT(MdBdyLy(A.Md), A.Nm)
End Function
Function MthLyCxt(MthLy$()) As String()
Stop '
'MthLyCxt = CxtIx(MthLy, FTNo(1, Sz(MthLy)))
End Function

