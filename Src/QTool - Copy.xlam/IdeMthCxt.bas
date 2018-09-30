Attribute VB_Name = "IdeMthCxt"
Option Explicit
Function MthCxtFT(A As Mth) As FTNo()
MthCxtFT = SrcMthCxtFT(MdBdyLy(A.Md), A.Nm)
End Function

Function MthLyCxt(MthLy$()) As String()
MthLyCxt = CxtIx(MthLy, FTNo(1, Sz(MthLy)))
End Function

Private Function CxtIx(Src$(), MthIx As FTIx) As FTIx
'Src -> X:MthFmno -> MthCxtFTNo
With MthIx
    Dim Ix%
    For Ix = .Fmix To .Toix
        If Not LasChr(Src(Ix)) = "_" Then
            Ix = Ix + 1
            Exit For
        End If
    Next
    Set CxtIx = FTIx(Ix, .Toix - 1)
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
