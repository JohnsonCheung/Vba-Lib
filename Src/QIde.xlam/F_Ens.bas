Attribute VB_Name = "F_Ens"
Option Explicit
Private Const C_FldLvs$ = "Lx Src Som NewLin Md"

Function PjZZDrs(A As VBProject) As Drs
Dim Dry(), I
'For Each I In PjMdAy(A)
    PushAy Dry, MdDry(CvMd(I))
'Next
PjZZDrs = NewDrs(C_FldLvs, Dry)
End Function

Sub MdEnsZZAsPrivate(A As CodeModule)
DrsEns MdDrs(A)
End Sub

Sub PjEnsZZAsPrivate(A As VBProject)
DrsEns PjDrs(A)
End Sub

Private Sub DrsEns(A As Drs)
If AyIsEmp(A.Dry) Then Exit Sub
Dim Lx%, Md As CodeModule, NewL$, OldL$, Som As Boolean, Dr
Dim X%():         X = FnyIxAy(A.Fny, "Lx Src Som NewLin Md")
Dim Drs As Drs: Drs = MdDrs(Md)

For Each Dr In A.Dry
    DrIxAy_Asg Dr, X, _
        Lx, OldL, Som, NewL, Md
    If Not Som Then Stop
    If Md.Lines(Lx, 1) <> OldL Then Stop
    MdRplLin Md, Lx, NewL    '<-- Update
    Debug.Print MdNm(Md), Lx, OldL, NewL
Next
End Sub

Private Function MdDrs(A As CodeModule) As Drs
Dim Dry(): Dry = MdDry(A)
MdDrs = NewDrs(C_FldLvs, Dry)
End Function

Private Function MdDry(A As CodeModule) As Variant()
MdDry = SrcDry(MdSrc(A), MdNm(A))
End Function

Private Function PjDrs(A As VBProject) As Drs
Dim ODry()
Dim I
For Each I In Pjx(A).MdAy
    PushAy ODry, MdDry(CvMd(I))
Next
PjDrs = NewDrs(C_FldLvs, ODry)
End Function

Private Function SrcDry(A$(), MdNm$) As Variant()
Dim Src1$()
Dim Som() As Boolean
Dim NewL$()
Dim J%
Dim Dry1()

Dim M() As StrOpt
Dim Lx&(): Lx = SrcLxAy(A)
    M = SrcLxAy_EnsPrivate(A, Lx)
    If StrOptAy_HasNone(M) Then Stop

Dim Dry()
    Src1 = AyWh_ByIxAy(A, Lx)
    For J = 0 To UB(Src1)
        With MthLin_EnsPrivate(Src1(J))
            Push Som, .Som
            Push NewL, .Str
        End With
    Next
    Dry1 = AyZipAp(Lx, Src1, Som, NewL)
    Dry = DryAddConstCol(Dry1, MdNm)
SrcDry = Dry
End Function

Private Function SrcEnsZZAsPrivateDrs(A$(), LxAy&()) As Drs

End Function

Private Function SrcLxAy(A$()) As Long()
Dim L$, Lx%, J%, MthNm$
Dim O1&(): O1 = SrcMthLxAy(A)
Dim O&()
    For J = 0 To UB(O1)
        Lx = O1(J)
        L = A(Lx)
        If HasPfx(L, "Private ") Then GoTo Nxt
        MthNm = SrcLin_MthNm(L)
        If Not HasPfx(MthNm, "ZZ") Then GoTo Nxt
        Push O, O1(J)
Nxt:
    Next
SrcLxAy = O
End Function

Private Function SrcLxAy_EnsPrivate(A$(), LxAy&()) As StrOpt()
Dim U%: U = UB(LxAy)
If U = -1 Then Exit Function
Dim O() As StrOpt
    ReDim O(U)
    Dim J%
    For J = 0 To U
        O(J) = MthLin_EnsPrivate(A(LxAy(J)))
    Next
SrcLxAy_EnsPrivate = O
End Function

Private Sub ZZZPjEnsZZAsPrivate()
PjEnsZZAsPrivate CurPj
End Sub

Private Sub ZZ_MdDrs()
'DrsBrw MdDrs(Md("IdeFeature_EnsZZ_AsPrivate"))
End Sub

Private Sub ZZ_PjDrs()
DrsBrw PjDrs(CurPj)
End Sub
