Attribute VB_Name = "A_1"
Option Explicit
Private B_MdNm$
Private Fso As New Scripting.FileSystemObject
Const SrcPjNm$ = "Lib_XX"
Const TarPjNm$ = "VbLib"
Private Sub ZZ_MovMd()
MovMd "FmTo"
End Sub
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
Private Sub ZRmvFst4Lines(Ft$)
Dim A$: A = Fso.GetFile(Ft).OpenAsTextStream.ReadAll
Dim B$: B = Left(A, 55)
Dim C$: C = Mid(A, 56)
Dim B1$: B1 = Replace("VERSION 1.0 CLASS|BEGIN|  MultiUse = -1  'True|END|", "|", vbCrLf)
If B <> B1 Then Stop
Fso.CreateTextFile(Ft, True).Write C
End Sub
Private Function ZMdTy() As vbext_ComponentType
ZMdTy = ZSrcCmp.Type
End Function
Private Function ZSrcCmp() As VBComponent
Set ZSrcCmp = ZSrcPj.VBComponents(B_MdNm)
End Function
Private Sub Ass(B As Boolean)
If Not B Then Stop
End Sub
Private Function ZMd_NotExist_InTar() As Boolean
Dim I, Cmp As VBComponent
For Each I In ZTarPj.VBComponents
    Set Cmp = I
    If Cmp.Name = B_MdNm Then Exit Function
Next
ZMd_NotExist_InTar = True
End Function
Private Function ZTarPj() As VBProject
Set ZTarPj = Excel.Application.VBE.VBProjects(TarPjNm)
End Function
Private Function ZSrcPj() As VBProject
Set ZSrcPj = Excel.Application.VBE.VBProjects(SrcPjNm)
End Function
Private Function TmpFt$(Optional Fdr$, Optional Fnn$)
TmpFt = TmpFfn(".txt", Fdr, Fnn)
End Function

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
Private Sub PthEns(P$)
If Fso.FolderExists(P) Then Exit Sub
MkDir P
End Sub

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

