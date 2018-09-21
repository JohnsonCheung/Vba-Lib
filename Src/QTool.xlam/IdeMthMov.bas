Attribute VB_Name = "IdeMthMov"
Option Explicit
Function MthCpy(A As Mth, ToMd As CodeModule, Optional IsSilent As Boolean) As Boolean
If MdHasMth(ToMd, A.Nm) Then
    Debug.Print FmtQQ("MthCpy_ToMd: Fm-Mth(?) is Found in To-Md(?)", A.Nm, MdNm(ToMd))
    MthCpy = True
    Exit Function
End If
If ObjPtr(A.Md) = ObjPtr(ToMd) Then
    Debug.Print FmtQQ("MthCpy: Fm-Mth-Md(?) cannot be To-Md(?)", MthMdNm(A), MdNm(ToMd))
    MthCpy = True
    Exit Function
End If
MdAppLines ToMd, MthLines(A)
If Not IsSilent Then
    Debug.Print FmtQQ("MthCpy: Mth(?) is copied ToMd(?)", MthDNm(A), MdDNm(ToMd))
End If
End Function

Sub MthCpyToPj(A As Mth, ToPj As VBProject, Optional IsSilent As Boolean)
Dim ToMdNm$: ToMdNm = MthProperMdNm(A.Nm)
Dim ToMd As CodeModule: Set ToMd = PjMd(ToPj, ToMdNm)
MthCpy A, ToMd
End Sub

Sub MthAyMov(A() As Mth, ToMd As CodeModule)
AyDoXP A, "MthMov", ToMd
End Sub
Sub MovMth(MthPatn$, ToMdNm$)
CurMdMovMth MthPatn, Md(ToMdNm)
End Sub
Sub CurMdMovMth(MthPatn$, ToMd As CodeModule)
MdMovMth CurMd, MthPatn, ToMd
End Sub

Sub MdMovMth(A As CodeModule, MthPatn$, ToMd As CodeModule)
Dim MthNy$(), M
MthNy = AyWhPatn(MdMthNy(A, "Pub"), MthPatn)
For Each M In AyNz(MthNy)
    MthMov Mth(A, M), ToMd
Next
End Sub

Sub MthMov(A As Mth, ToMd As CodeModule)
If MthCpy(A, ToMd, IsSilent:=True) Then Exit Sub
MthRmv A, IsSilent:=True
Debug.Print FmtQQ("MthMov: Mth(?) is moved to Md(?)", MthDNm(A), MdDNm(ToMd))
End Sub

Sub MthMovToProperMd(A As Mth)
If MdCmpTy(A.Md) <> vbext_ct_StdModule Then
    Debug.Print FmtQQ("MthMovToProperMd: Md(?) in not in StdMd", MthDNm(A))
    Exit Sub
End If
If Not IsPfx(A.Nm, "ZZ_") Then
    If Not MthIsPub(A) Then
        Debug.Print FmtQQ("MdMovToProperMd: Mth(?) is not public", MthDNm(A))
        Exit Sub
    End If
End If
MthMov A, MthProperMd(A)
End Sub

Sub CurMthMov(ToMd$)
MthMov CurMth, Md(ToMd)
End Sub
