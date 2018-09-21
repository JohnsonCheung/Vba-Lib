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

Private Function IsPrvZZDashMthLin(A) As Boolean
Dim L$, M$, T$
AyAsg ShiftMdy(A), M, L: If M <> "Private" Then Exit Function
AyAsg ShiftMthTy(L), T, L: If T = "" Then Exit Function
If Left(L, 3) <> "ZZ_" Then Exit Function
IsPrvZZDashMthLin = True
End Function

Private Function IsPubZZDashMthLin(A) As Boolean
Dim L$, M$, T$
AyAsg ShiftMdy(A), M, L: If M <> "Public" And M <> "" Then Exit Function
AyAsg ShiftMthTy(L), T, L: If T = "" Then Exit Function
If Left(L, 3) <> "ZZ_" Then Exit Function
IsPubZZDashMthLin = True
End Function

Function IsPubZDashMthLin(A) As Boolean
Dim L$, M$, T$
AyAsg ShiftMdy(A), M, L: If M <> "Public" And M <> "" Then Exit Function
AyAsg ShiftMthTy(L), T, L: If T = "" Then Exit Function
If IsPfx(L, "Z__Tst()") Then Exit Function
If Not IsPfx(L, "Z_") Then Exit Function
IsPubZDashMthLin = True
End Function

Sub MdEnsZ3DMthAsPrivate(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZ3DMthAsPrivate: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If IsPubZDashMthLin(L) Then
        By = MthLin_EnsPrivate(L)
        Debug.Print FmtQQ("MdEnsZ3DMthAsPrivate Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub
Sub PjEnsZDashMthAsPrv(A As VBProject)
ItrDo PjCdMdAy(A), "MdEnsZ3DMthAsPrivate"
End Sub

Private Sub MdEnsZZDashPrvMthAsPub(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If IsPrvZZDashMthLin(L) Then
        By = MthLin_EnsPublic(L)
        Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPub: Md(?) Lin(?) is change to Public: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub

Private Sub MdEnsZZDashPubMthAsPrv(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPubMthAsPrv Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If IsPubZZDashMthLin(L) Then
        Debug.Print L
        By = MthLin_EnsPrivate(L)
        Debug.Print FmtQQ("MdEnsZZDashPubMthAsPrv: Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        'A.ReplaceLine J, By
    End If
Next
End Sub

Sub MthEnsPrv(A As Mth)
Dim F%(): F = MthLnoAy(A)
Dim F1%(), N$
Dim L2$, J%, L1$, L$
N = MthDNm(A)
For J = 0 To UB(F)
    L = F(J)
    L1 = A.Md.Lines(L, 1)
    L2 = MthLin_EnsPrivate(L1)
    If L1 <> L2 Then
        Debug.Print FmtQQ("MthEnsPub: Md(?) Lin(?) RplBy(?)", N, L, L2)
        A.Md.ReplaceLine L, L2
    End If
Next
End Sub

Sub MthEnsPub(A As Mth)
Dim F%(): F = MthLnoAy(A)
Dim F1%(), N$
Dim L2$, J%, L1$, L$
N = MthDNm(A)
For J = 0 To UB(F)
    L = F(J)
    L1 = A.Md.Lines(L, 1)
    L2 = MthLin_EnsPublic(L1)
    If L1 <> L2 Then
        Debug.Print FmtQQ("MthEnsPub: Md(?) Lin(?) RplBy(?)", N, L, L2)
        A.Md.ReplaceLine L, L2
    End If
Next
End Sub

Private Function MthLin_EnsPrivate$(A)
Dim L$: L = A
ShiftMdy L
MthLin_EnsPrivate = "Private " & L
End Function

Private Function MthLin_EnsPublic$(A)
MthLin_EnsPublic = RmvMdy(A)
End Function

Sub PjEnsZZDashPubMthAsPrivate(A As VBProject)
AyDo PjCdMdAy(A), "MdEnsZZDashPubMthAsPrv"
End Sub

Sub VbeEnsZZDashPubMthAsPrivate(A As Vbe)
AyDo VbePjAy(A), "PjEnsZZDashPubMthAsPrivate"
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

Private Sub Z_MthEnsPub()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "YYA")
MthEnsPrv M: Debug.Assert MthLin(M) = "Private Property Get ZZA()"
MthEnsPub M:  Debug.Assert MthLin(M) = "Property Get ZZA()"
End Sub
