Attribute VB_Name = "IdeMthMdy"
Option Explicit

Sub Ens_Md_Z3DMthAsPrivate()
MdEnsZ3DMthAsPrivate CurMd
End Sub
Sub Ens_Pj_Z3DMthAsPrivate()
PjEnsZDashMthAsPrv CurPj
End Sub

Sub Ens_Md_ZZDashPrvMthAsPublic()
MdEnsZZDashPrvMthAsPub CurMd
End Sub

Sub Ens_Md_ZZDashPubMthAsPrivate()
MdEnsZZDashPubMthAsPrv CurMd
End Sub

Sub Ens_Pj_ZZDashPubMthAsPrivate()
PjEnsZZDashPubMthAsPrivate CurPj
End Sub

Sub Ens_Vbe_ZZDashPubMthAsPrivate()
VbeEnsZZDashPubMthAsPrivate CurVbe
End Sub

Function IsPrvZZDashMthLin(A) As Boolean
Dim L$, M$, T$
Stop '
'AyAsg ShfMdy(A), M, L: If M <> "Private" Then Exit Function
'AyAsg ShfMthTy(L), T, L: If T = "" Then Exit Function
'If Left(L, 3) <> "ZZ_" Then Exit Function
IsPrvZZDashMthLin = True
End Function

Function IsPubZZDashMthLin(A) As Boolean
Dim L$, M$, T$
Stop '
'AyAsg ShfMdy(A), M, L: If M <> "Public" And M <> "" Then Exit Function
'AyAsg ShfMthTy(L), T, L: If T = "" Then Exit Function
'If Left(L, 3) <> "ZZ_" Then Exit Function
'IsPubZZDashMthLin = True
End Function

Function IsPubZDashMthLin(A) As Boolean
Dim L$, M$, T$
Stop '
'AyAsg ShfMdy(A), M, L: If M <> "Public" And M <> "" Then Exit Function
'AyAsg ShfMthTy(L), T, L: If T = "" Then Exit Function
'If IsPfx(L, "Z__Tst()") Then Exit Function
'If Not IsPfx(L, "Z_") Then Exit Function
'IsPubZDashMthLin = True
End Function

Sub PjEnsZDashMthAsPrv(A As VBProject)
ItrDo PjMdAy(A), "MdEnsZ3DMthAsPrivate"
End Sub







Sub PjEnsZZDashPubMthAsPrivate(A As VBProject)
AyDo PjMdAy(A), "MdEnsZZDashPubMthAsPrv"
End Sub


Private Sub ZZ_LinIsNonPrivateZMthLin()
Dim S$(): S = MdSrc(CurMd)
Dim N$(): N = AyWhPred(S, "LinIsNonPrivateZMthLin")
Brw N
End Sub

Private Sub ZZ_IsPrvZZDashMthLin()
Dim L
For Each L In CurSrc
    If IsPrvZZDashMthLin(L) Then
        Debug.Print L
    End If
Next
End Sub

Private Sub ZZ_IsPubZZDashMthLin()
Dim L
For Each L In CurSrc
    If IsPubZZDashMthLin(L) Then
        Debug.Print L
    End If
Next
End Sub

