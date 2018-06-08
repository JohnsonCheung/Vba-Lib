Attribute VB_Name = "A_1"
Option Explicit
Private B_MdNm$
Private Sub ZZ_MovMd()
MovMd "M_Vb"
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
Dim TarCmp As VBComponent
    Set TarCmp = ZTarPj.VBComponents.Add(ZMdTy)
    TarCmp.CodeModule.AddFromFile TmpFil
ZSrcPj.VBComponents.Remove SrcCmp
Kill TmpFil
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
Set ZTarPj = Excel.Application.Vbe.VBProjects("VbLib")
End Function
Private Function ZSrcPj() As VBProject
Set ZSrcPj = Excel.Application.Vbe.VBProjects("Lib_XX")
End Function

