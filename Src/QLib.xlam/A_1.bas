Attribute VB_Name = "A_1"
Option Explicit
Private B_MdNm$
Private Fso As New Scripting.FileSystemObject
Const SrcPjNm$ = "Lib_XX"
Const TarPjNm$ = "VbLib"

Sub MovMd(MdNm$)
B_MdNm = MdNm
'Move the MdNm in SrcPj-(Lib_XX) to TarPj-(VbLib)
Ass ZMd_NotExist_InTar
Dim SrcCmp As VBComponent
Dim TmpFil$
    TmpFil = TmpFfn(".txt")
    Set SrcCmp = ZSrcCmp
    SrcCmp.Export TmpFil
    If SrcCmp.Type = vbext_ct_ClassModule Then
        ZRmvFst4Lines TmpFil
    End If
Dim TarCmp As VBComponent
    Set TarCmp = ZTarPj.VBComponents.Add(ZMdTy)
    TarCmp.CodeModule.AddFromFile TmpFil
ZSrcPj.VBComponents.Remove SrcCmp
Kill TmpFil
End Sub

Sub MovMdLik(MdLikNm$)
Dim Ny$(): Ny = ZMdNyLik(MdLikNm)
Dim I
For Each I In Ny
    MovMd CStr(I)
Next
End Sub

Function TmpFfn(Ext$, Optional Fdr$, Optional Fnn0$)
Dim Fnn$
    Fnn = IIf(Fnn0 = "", TmpNm, Fnn0)
TmpFfn = TmpPth(Fdr) & Fnn & Ext
End Function

Function TmpNm$()
Static X&
TmpNm = "T" & Format(Now(), "YYYYMMDD_HHMMSS") & "_" & X
X = X + 1
End Function

Function TmpPth$(Optional Fdr$)
Dim X$
   If Fdr <> "" Then
       X = Fdr & "\"
   End If
Dim O$
   O = TmpPthFix & X:   PthEns O
   O = O & TmpNm & "\": PthEns O
   PthEns O
TmpPth = O
End Function

Function TmpPthFix$()
Static X$
If X = "" Then X = Fso.GetSpecialFolder(TemporaryFolder) & "\"
TmpPthFix = X
End Function

Private Sub Ass(B As Boolean)
If Not B Then Stop
End Sub

Private Sub PthEns(P$)
If Fso.FolderExists(P) Then Exit Sub
MkDir P
End Sub

Private Sub Push(OAy, M)
Dim N&: N = Sz(OAy)
ReDim Preserve OAy(N)
If IsObject(M) Then
    Set OAy(N) = M
Else
    OAy(N) = M
End If
End Sub

Private Property Get Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Property

Private Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

Private Property Get ZMdNyLik(MdLikNm$)
Dim Cmp As VBComponent, O$()
For Each Cmp In ZSrcPj.VBComponents
    If Cmp.Name Like MdLikNm Then
        Push O, Cmp.Name
    End If
Next
ZMdNyLik = O
End Property

Private Function ZMdTy() As vbext_ComponentType
ZMdTy = ZSrcCmp.Type
End Function

Private Function ZMd_NotExist_InTar() As Boolean
Dim I, Cmp As VBComponent
For Each I In ZTarPj.VBComponents
    Set Cmp = I
    If Cmp.Name = B_MdNm Then Exit Function
Next
ZMd_NotExist_InTar = True
End Function

Private Sub ZRmvFst4Lines(Ft$)
Dim A$: A = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
Dim B$: B = Left(A, 55)
Dim C$: C = Mid(A, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If B <> B1 Then Stop
Fso.CreateTextFile(Ft, True).Write C
End Sub

Private Function ZSrcCmp() As VBComponent
Set ZSrcCmp = ZSrcPj.VBComponents(B_MdNm)
End Function

Private Function ZSrcPj() As VBProject
Set ZSrcPj = Excel.Application.Vbe.VBProjects(SrcPjNm)
End Function

Private Function ZTarPj() As VBProject
Set ZTarPj = Excel.Application.Vbe.VBProjects(TarPjNm)
End Function

Private Sub ZZ_MovMd()
MovMdLik "Vb*"
End Sub
