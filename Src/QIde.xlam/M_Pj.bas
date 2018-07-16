Attribute VB_Name = "M_Pj"
Option Explicit
Property Get PjEnsCmp(A As VBProject, MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
If PjHasCmp(A, MdNm) Then
    PjCrtCmp A, MdNm, Ty
End If
Set PjEnsCmp = PjMd(A, MdNm)
End Property

Property Get PjMd(A As VBProject, MdNm) As CodeModule
Set PjMd = A.VBComponents(MdNm).CodeModule
End Property

Property Get PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.Filename
End Property

Property Get PjIsFxa(A As VBProject) As Boolean
PjIsFxa = LCase(FfnExt(PjFfn(A))) = ".xlam"
End Property

Property Get PjIsUsrLib(A As VBProject) As Boolean
PjIsUsrLib = PjIsFxa(A)
End Property

Sub PjAddRf(A As VBProject, RfNm, PjFfn)
If PjHasRf(A, RfNm) Then Exit Sub
A.References.AddFromFile PjFfn
End Sub

Sub PjCpyToSrcPth(A As VBProject)
FfnCpyToPth A.Filename, PjSrcPth(A), OvrWrt:=True
End Sub

Function PjCrtFxa(FxaNm$) As Excel.Application
Stop '
'Dim F$: F = FfnPth(A.Filename) & FxaNm & ".xlam"
'Set CrtFxa = Fxa(F).Crt
End Function

Function PjDltMd(A As VBProject, MdNm$) As Boolean
If Not PjHasCmp(A, MdNm) Then Exit Function
A.VBComponents.Remove A.VBComponents(MdNm)
PjDltMd = True
End Function

Sub PjGoMdNm(A As VBProject, MdNm$, Optional ClsOth As Boolean)
If ClsOth Then WinClsCd
Md(MdNm).CodePane.Show
End Sub

Sub PjImpSrcFfn(A As VBProject, SrcFfn)
A.VBComponents.Import SrcFfn
End Sub

Function Pj(PjNm$) As VBProject
Set Pj = CurVbe.VBProjects(PjNm)
End Function

Sub PjCrtMd(A As VBProject, MdNm$)
PjCrtCmp A, MdNm, vbext_ct_StdModule
End Sub

Function PjClsAy(A As VBProject, Optional NmPatn$ = ".") As CodeModule()
PjClsAy = PjMdAy(A, NmPatn, CmpTyAyOfCls)
End Function

Function PjClsNy(A As VBProject, Optional NmPatn$ = ".") As String()
PjClsNy = OyNy(PjClsAy(A, NmPatn))
End Function

Sub PjCpyToSrc(A As VBProject)
FfnCpyToPth A.Filename, PjSrcPth(A), OvrWrt:=True
End Sub

Sub PjCrtCmp(A As VBProject, Nm$, Ty As vbext_ComponentType)
If PjHasCmp(A, Nm) Then Stop
Dim O As VBComponent
Set O = A.VBComponents.Add(Ty)
O.Name = Nm
O.CodeModule.InsertLines 1, "Option Explicit"
End Sub

Sub PjExp(A As VBProject)
PjExpSrc A
PjExpRf A
End Sub

Sub PjExpRf(A As VBProject)
Ass Not PjIsUnderSrcPth(A)
AyWrt PjRfLy(A), PjRfCfgFfn(A)
End Sub

Sub PjExpSrc(A As VBProject)
PjCpyToSrc A
PthClrFil PjSrcPth(A)
Dim Md As CodeModule, I
For Each I In PjMbrAy(A)
    Set Md = I
    MdExp Md
Next
End Sub

Function PjHasCmp(A As VBProject, CmpNm$) As Boolean
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If CmpNm = Cmp.Name Then
        PjHasCmp = True
        Exit Function
    End If
Next
End Function

Function PjHasRf(A As VBProject, RfNm)
Dim RF As VBIDE.Reference
For Each RF In A.References
    If RF.Name = RfNm Then PjHasRf = True: Exit Function
Next
End Function

Sub PjImpRf(A As VBProject, RfCfgPth$)
Dim B As Dictionary: Set B = FtDic(RfCfgPth & "PjRf.Cfg")
Dim K
For Each K In B.Keys
    PjAddRf A, K, B(K)
Next
End Sub

Function PjIsUnderSrcPth(A As VBProject) As Boolean
Dim B$: B = PjPth(A)
If PthFdr(B) = "Src" Then Stop
End Function

Function PjMbrAy(A As VBProject, Optional NmPatn$ = ".", Optional CmpTyAy0) As CodeModule()
Dim Ay() As vbext_ComponentType
Ay = DftCmpTyAy(CmpTyAy0)

Dim Ay() As vbext_ComponentType
PjMbrAy = PjCmpAy(A, Ay, NmPatn)
End Function

Function PjMbrNy(A As VBProject, Optional NmPatn$ = ".") As String()
PjMbrNy = OyPrp(PjMdAy(A, NmPatn), "Name", EmpSy)
End Function

Function PjMdAy(A As VBProject, Optional NmPatn$ = ".") As CodeModule()
PjMdAy = ZPjMbrAy(A, Ay, NmPatn, CmpTyAy0)
End Function

Function PjMdAyOfStd(A As VBProject, Optional NmPatn$ = ".") As CodeModule()
PjMdAyOfStd = PjMdAy(A, NmPatn, CmpTyAyOfStd)
End Function

Function PjMdNy(A As VBProject, Optional NmPatn$ = ".", Optional CmpTyAy0) As String()
PjMdNy = OyPrp(PjMdAy(A, NmPatn, CmpTyAy0), "Name", EmpSy)
End Function

Function PjMdNyOfStd(A As VBProject, Optional NmPatn$ = ".") As String()
PjMdNyOfStd = PjMdNy(A, NmPatn, CmpTyAyOfStd)
End Function

Function PjMthDrs(A As VBProject, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
Dim Fny$()
    Fny = FnyOfMthDrs(WithBdyLy, WithBdyLines)
PjMthDrs.Fny = Fny
PjMthDrs.Dry = PjMthDry(A, WithBdyLy, WithBdyLines)
End Function

Function PjMthDry(A As VBProject, Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In PjMdAy(A)
        Set Md = I
        PushAy Dry, MdMthDry(Md, WithBdyLy, WithBdyLines)
    Next
PjMthDry = Dry
End Function

Function PjMthLinDry(A As VBProject) As Variant()
Dim I, Md As CodeModule, O()
For Each I In PjMbrAy(A)
    Set Md = I
    Dim N$: N = MdNm(Md)
    Dim Ay$(): Ay = MdMthLinAy(Md)
    Dim Dry(): Dry = ConstAy_ConstValDry(N, Ay)
    PushAy O, Dry
Next
PjMthLinDry = O
End Function

Function PjMthNy(A As VBProject, Optional CmpTyAy0, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = ".") As String()
Dim CmpTyAy() As vbext_ComponentType
    CmpTyAy = DftCmpTyAy(CmpTyAy0)
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(PjMdNy(A, MdNmPatn))
If AyIsEmp(Ny) Then Exit Function
For Each I In Ny
    Set M = PjMd(A, CStr(I))
    PushAy O, AyAddPfx(MdMthNy(M, MthNmPatn), MdNm(M) & Sep)
Next
PjMthNy = O
End Function

Function PjNm$(A As VBProject)
PjNm = A.Name
End Function

Function PjPatnLy(A As VBProject, Patn$) As String()
Dim I, Md As CodeModule, O$()
For Each I In PjMdAy(A)
   Set Md = I
   PushAy O, MdPatnLy(Md, Patn)
Next
PjPatnLy = O
End Function

Function PjPth$(A As VBProject)
PjPth = FfnPth(A.Filename)
End Function

Function PjReadRfCfg(A As VBProject) As String()
Const CSub$ = "PjReadRfCfg"
Dim B$: B = PjRfCfgFfn(A)
If Not FfnIsExist(B) Then Er CSub, "{Pj-Rf-Cfg-Fil} not found", B
PjReadRfCfg = FtLy(B)
End Function

Sub PjRenMdByPfx(A As VBProject, FmMdPfx$, ToMdPfx$)
Dim DftNy$()
Dim Ny$()
    Ny = PjMdNy(A, "^" & FmMdPfx)
    DftNy = AyMapAsgSy(Ny, "RplPfx", FmMdPfx, ToMdPfx)
Dim MdAy() As CodeModule
    Dim MdNm
    Dim Md As CodeModule
    For Each MdNm In Ny
        Set Md = PjMd(A, CStr(MdNm))
        PushObj MdAy, Md
    Next
Dim I%, U%
    For I = 0 To UB(DftNy)
        MdRen MdAy(I), DftNy(I)
    Next
End Sub

Function PjRfAy(A As VBProject) As VBIDE.Reference()
Dim RF As VBIDE.Reference, O() As VBIDE.Reference
For Each RF In A.References
    Push O, RF
Next
PjRfAy = O
End Function

Sub PjRfBrw(A As VBProject)
aybrw PjRfLy(A)
End Sub

Function PjRfCfgFfn$(A As VBProject)
PjRfCfgFfn = PjSrcPth(A) & "PjRf.Cfg"
End Function

Sub PjRfDmp(A As VBProject)
AyDmp PjRfLy(A)
End Sub

Function PjRfLy(A As VBProject) As String()
Dim RfAy() As VBIDE.Reference
    RfAy = PjRfAy(A)
Dim O$()
Dim Ny$(): Ny = OyPrpSy(RfAy, "Name")
Ny = AyAlignL(Ny)
Dim J%
For J = 0 To UB(Ny)
    Push O, Ny(J) & " " & RfPth(RfAy(J))
Next
PjRfLy = O
End Function

Sub PjRmvMdNmPfx(A As VBProject, Pfx$)
Dim I, Md As CodeModule
For Each I In PjMdAy(A, "^" & Pfx)
    Set Md = I
    MdRmvNmPfx Md, Pfx
Next
End Sub

Function PjSrcPth$(A As VBProject)
Dim Ffn$: Ffn = PjFfn(A)
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(A.Filename) & "Src\": PthEns O
O = O & Fn & "\":                PthEns O
PjSrcPth = O
End Function

Sub PjSrcPthBrw(A As VBProject)
PthBrw PjSrcPth(A)
End Sub

Function PjTyNy(A As VBProject, Optional TyNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = vbTab) As String()
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(PjMdNy(A, MdNmPatn))
If AyIsEmp(Ny) Then Exit Function
For Each I In Ny
    Set M = PjMd(A, CStr(I))
    PushAy O, AyAddPfx(MdTyNy(M, TyNmPatn), MdNm(M) & Sep)
Next
PjTyNy = O
End Function

Sub PjRmvOptCmpDbLin()
Dim I
For Each I In MdAy
   'MdRmvOptCmpDb CvMd(I)
Next
End Sub

Sub PjSrt(A As VBProject)
Dim M As CodeModule, I, Ay() As CodeModule
If AyIsEmp(Ay) Then Exit Sub
For Each I In Ay
    MdSrt CvMd(I)
Next
End Sub

Private Sub PjCurPjx__Tst()
Ass CurPj.Name = "lib1"
End Sub

Sub PjClsNy__Tst()
AyDmp PjClsNy(CurPj)
End Sub

Private Sub PjMdAy__Tst()
Dim O() As CodeModule
O = PjMdAy(CurPj)
Dim I, Md As CodeModule
For Each I In O
    Set Md = I
    Debug.Print MdNm(Md)
Next
End Sub

Sub PjMdNy__Tst()
AyDmp PjMdNy(CurPj)
End Sub

Private Sub PjMthDrs__Tst()
Dim Drs As Drs
Drs = PjMthDrs(CurPj, WithBdyLines:=True)
WsVis DrsWs(Drs, PjNm(CurPj))
End Sub

Private Sub PjMthLinDry__Tst()
Dim A(): A = PjMthLinDry(CurPj)
Stop
End Sub

Private Sub PjRenMdByPfx__Tst()
PjRenMdByPfx CurPj, "A_", ""
End Sub

Private Property Get CmpTyAyOfCls() As vbext_ComponentType()
Dim T(0) As vbext_ComponentType
T(0) = vbext_ct_ClassModule
CmpTyAyOfCls = T
End Property

Private Property Get CmpTyAyOfStd() As vbext_ComponentType()
Dim Ay(0) As vbext_ComponentType
Ay(0) = vbext_ct_StdModule
CmpTyAyOfStd = Ay
End Property

Private Property Get ZMbrAy(MbrTyAy() As vbext_ComponentType, Optional NmPatn$ = ".") As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
Dim NmRe As Re: Set NmRe = Re(NmPatn)
Dim Sel As Boolean: Sel = Sz(MbrTyAy) > 0
For Each Cmp In A.VBComponents
    If Not NmRe.Tst(Cmp.Name) Then GoTo X
    If Sel Then
        If AyHas(MbrTyAy, Cmp.Type) Then
            PushObj O, Cmp.CodeModule
        End If
    Else
        PushObj O, Cmp.CodeModule
    End If
X:
Next
ZMbrAy = O
End Property

Private Function PjZPjMbrAy(A As VBProject, MbrTyAy() As vbext_ComponentType, Optional NmPatn$ = ".") As CodeModule()
Dim O() As CodeModule
Dim Cmp As VBComponent
Dim Sel As Boolean: Sel = Sz(MbrTyAy) > 0
For Each Cmp In A.VBComponents
    If Not ReTst(Cmp.Name, NmPatn) Then GoTo X
    If Sel Then
        If AyHas(MbrTyAy, Cmp.Type) Then
            PushObj O, Cmp.CodeModule
        End If
    Else
        PushObj O, Cmp.CodeModule
    End If
X:
Next
ZPjMbrAy = O
End Function

Private Sub PjZZ_MthNy()
Stop
'AyBrw Pjx(CurPj).MthNy
End Sub

Private Sub PjZZ_PjMthNy()
aybrw PjMthNy(CurPjx)
End Sub
