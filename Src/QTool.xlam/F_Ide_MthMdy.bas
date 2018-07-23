Attribute VB_Name = "F_Ide_MthMdy"
Option Explicit

Sub Ens_Md_ZZDashPrvMthAsPublic()
MdEnsZZDashPrvMthAsPublic CurMd
End Sub

Sub Ens_Md_ZZDashPubMthAsPrivate()
MdEnsZZDashPubMthAsPrivate CurMd
End Sub

Sub Ens_Pj_ZZDashPubMthAsPrivate()
PjEnsZZDashPubMthAsPrivate CurPj
End Sub

Sub Ens_Vbe_ZZDashPubMthAsPrivate()
VbeEnsZZDashPubMthAsPrivate CurVbe
End Sub

Private Function LinIsPrivateZZDashMthLin(A) As Boolean
Dim L$: L = A
Dim M$: M = LinShiftMdy(L): If M <> "Private" Then Exit Function
Dim T$: T = LinShiftMthTy(L): If T = "" Then Exit Function
If FstChr(L) <> "ZZ_" Then Exit Function
LinIsPrivateZZDashMthLin = True
End Function

Private Function LinIsPublicZZDashMthLin(A) As Boolean
Dim L$: L = A
Dim M$: M = LinShiftMdy(L): If M <> "" Then Exit Function
Dim T$: T = LinShiftMthTy(L): If T = "" Then Exit Function
If FstChr(L) <> "ZZ_" Then Exit Function
If Not AscIsUCase(Asc(Mid(L, 2, 1))) Then Exit Function
LinIsPublicZZDashMthLin = True
End Function

Function LinIsZZDashPrvMth(A) As Boolean
Dim L$: L = A
Dim M$: M = LinShiftMdy(L): If M <> "Private" Then Exit Function
Dim T$: T = LinShiftMthTy(L): If T = "" Then Exit Function
If Not IsPfx(L, "ZZ_") Then Exit Function
LinIsZZDashPrvMth = True
End Function

Function LinIsZZDashPubMth(A) As Boolean
Dim L$: L = A
Dim M$: M = LinShiftMdy(L): If M <> "" And M <> "Public" Then Exit Function
Dim T$: T = LinShiftMthTy(L): If T = "" Then Exit Function
If Not IsPfx(L, "ZZ_") Then Exit Function
LinIsZZDashPubMth = True
End Function

Private Sub MdEnsZZDashPrvMthAsPublic(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPublic: Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If LinIsZZDashPrvMth(L) Then
        By = MthLin_EnsPublic(L)
        Debug.Print FmtQQ("MdEnsZZDashPrvMthAsPublic: Md(?) Lin(?) is change to Public: [?]", DNm, J, By)
        A.ReplaceLine J, By
    End If
Next
End Sub

Private Sub MdEnsZZDashPubMthAsPrivate(A As CodeModule)
Dim DNm$: DNm = MdDNm(A)
If DNm = "QTool.F_Ide_MthMdy" Then
    Debug.Print FmtQQ("MdEnsZZDashPubMthAsPrivate Md(?) is skipped", DNm)
    Exit Sub
End If
Dim J%, L$, By$
For J = 1 To A.CountOfLines
    L = A.Lines(J, 1)
    If LinIsZZDashPubMth(L) Then
        Debug.Print L
        By = MthLin_EnsPrivate(L)
        Debug.Print FmtQQ("MdEnsZZDashPubMthAsPrivate: Md(?) Lin(?) is change to Private: [?]", DNm, J, By)
        'A.ReplaceLine J, By
    End If
Next
End Sub

Sub MthEnsPrivate(A As Mth)
Dim F%(): F = MthFmLnoAy(A)
Dim F1%(), N$
Dim L2$, J%, L1$, L$
N = MthDNm(A)
For J = 0 To UB(F)
    L = F(J)
    L1 = A.Md.Lines(L, 1)
    L2 = MthLin_EnsPrivate(L1)
    If L1 <> L2 Then
        Debug.Print FmtQQ("MthEnsPublic: Md(?) Lin(?) RplBy(?)", N, L, L2)
        A.Md.ReplaceLine L, L2
    End If
Next
End Sub

Sub MthEnsPublic(A As Mth)
Dim F%(): F = MthFmLnoAy(A)
Dim F1%(), N$
Dim L2$, J%, L1$, L$
N = MthDNm(A)
For J = 0 To UB(F)
    L = F(J)
    L1 = A.Md.Lines(L, 1)
    L2 = MthLin_EnsPublic(L1)
    If L1 <> L2 Then
        Debug.Print FmtQQ("MthEnsPublic: Md(?) Lin(?) RplBy(?)", N, L, L2)
        A.Md.ReplaceLine L, L2
    End If
Next
End Sub

Private Function MthLin_EnsPrivate$(A)
Dim L$: L = A
LinShiftMdy L
MthLin_EnsPrivate = "Private " & L
End Function

Private Function MthLin_EnsPublic$(A)
MthLin_EnsPublic = LinRmvMdy(A)
End Function

Sub PjEnsZZDashPubMthAsPrivate(A As VBProject)
AyDo PjMbrAy(A), "MdEnsZZDashPubMthAsPrivate"
End Sub

Sub VbeEnsZZDashPubMthAsPrivate(A As Vbe)
AyDo VbePjAy(A), "PjEnsZZDashPubMthAsPrivate"
End Sub

Private Function ZZSrc() As String()
ZZSrc = MdSrc(Md("F_Ide_MthMdy"))
End Function

Private Sub ZZ_LinIsNonPrivateZMthLin()
Dim S$(): S = MdSrc(CurMd)
Dim N$(): N = AyWhPred(S, "LinIsNonPrivateZMthLin")
AyBrw N
End Sub

Private Sub ZZ_LinIsZZDashPrvMth()
Dim L
For Each L In ZZSrc
    If LinIsZZDashPrvMth(L) Then
        Debug.Print L
    End If
Next
End Sub

Private Sub ZZ_LinIsZZDashPubMth()
Dim L
For Each L In ZZSrc
    If LinIsZZDashPubMth(L) Then
        Debug.Print L
    End If
Next
End Sub

Private Sub ZZZ_MthEnsPublic()
Dim M As Mth: Set M = Mth(Md("ZZModule"), "ZZA")
MthEnsPrivate M: Debug.Assert MthLin(M) = "Private Property Get ZZA()"
MthEnsPublic M:  Debug.Assert MthLin(M) = "Property Get ZZA()"
End Sub
