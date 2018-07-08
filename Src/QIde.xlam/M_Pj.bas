Attribute VB_Name = "M_Pj"
Option Explicit

Function CrtFxa(FxaNm$) As Excel.Application
Stop '
'Dim F$: F = FfnPth(A.Filename) & FxaNm & ".xlam"
'Set CrtFxa = Fxa(F).Crt
End Function

Sub Sort()
Dim M As CodeModule, I
For Each I In MbrAy
    Set M = I
    MdSrt M
Next
End Sub

Property Get PjIsFxa(A As VBProject) As Boolean
PjIsFxa = LCase(FfnExt(PjFfn(A))) = ".xlam"
End Property

Property Get PjIsUsrLib(A As VBProject) As Boolean
PjIsUsrLin = PjIsFxa(A)
End Property
Sub PjCrtMd(A As VBProject, Nm$)
PjCrtCmp A, Nm, vbext_ct_StdModule
End Sub

Sub PjCrtCmp(A As VBProject, Nm$, Ty As vbext_ComponentType)
If PjHasCmp(A, Nm) Then Stop
Dim O As VBComponent
Set O = A.VBComponents.Add(vbext_ct_StdModule)
O.Name = Nm
O.CodeModule.InsertLines 1, "Option Explicit"
End Sub

Sub AddRf(RfNm, PjFfn)
If HasRf(RfNm) Then Exit Sub
A.References.AddFromFile PjFfn
End Sub

Property Get ClsAy(Optional NmPatn$ = ".") As CodeModule()
ClsAy = MdAy(NmPatn, CmpTyAyOfCls)
End Property

Property Get ClsNy(Optional NmPatn$ = ".") As String()
ClsNy = Oy(ClsAy(NmPatn)).Ny
End Property

Sub CpyToSrc()
Stop '
'FfnCpyToPth A.Filename, SrcPth, OvrWrt:=True
End Sub

Function CrtMd(Optional MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule, Optional RplBdy As Boolean) As CodeModule
If MdNm <> "" Then
    If HasMd(MdNm) Then
        Er "Pj.CrtMd", "Given {MdNm} exist", MdNm
        Exit Function
    End If
End If
Dim O As VBComponent: Set O = A.VBComponents.Add(Ty)
EnsMdExplicit O.Name
If MdNm <> "" Then
    O.Name = MdNm
End If
Set CrtMd = O.CodeModule
End Function

Function DltMd(MdNm$) As Boolean
If Not HasMd(MdNm) Then Exit Function
A.VBComponents.Remove A.VBComponents(MdNm)
DltMd = True
End Function

Sub Export()
Stop '
'FfnCpyToPth A.Filename, SrcPth, OvrWrt:=True
'PthClrFil SrcPth 'Clr SrcPth ---
'Oy(MbrAy).EachSub "Exp"
Dim I
For Each I In MbrAy
    'MdExp CvMd(I)  'Exp each md --
Next
AyWrt RfLy, RfCfgFfn 'Exp rf -----
End Sub

Sub ExpRf()
End Sub

Private Sub ExpSrc()
'CpyToSrc A
End Sub

Property Get Ffn$()
On Error Resume Next
Ffn = A.Filename
End Property

Sub ImpSrcFfn(SrcFfn)
A.VBComponents.Import SrcFfn
End Sub

Sub RmvOptCmpDbLin()
Dim I
For Each I In MdAy
   'MdRmvOptCmpDb CvMd(I)
Next
End Sub
Property Get HasMd(MdNm$) As Boolean
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If MdNm = Cmp.Name Then
        HasMd = True
        Exit Property
    End If
Next
End Property

Property Get HasMdNm(MdNm$) As Boolean
Dim I
Dim Cmp As VBComponent
For Each I In A.VBComponents
    Set Cmp = I
    If Cmp.Name = MdNm Then HasMdNm = True: Exit Property
Next
End Property

Property Get HasRf(RfNm)
Dim RF As VBIDE.Reference
For Each RF In A.References
    If RF.Name = RfNm Then HasRf = True: Exit Property
Next
End Property

Sub ImpRf(RfCfgPth$)
Stop '
'Dim B As Dictionary: Set B = FtDic(RfCfgPth & "PjRf.Cfg")
'Dim K
'For Each K In B.Keys
'    AddRf K, B(K)
'Next
End Sub

Property Get IsUnderSrcPth() As Boolean
Stop '
'If PthFdr(Pth) = "Src" Then Stop
End Property

Property Get MbrAy(Optional NmPatn$ = ".") As CodeModule()
Dim Ay() As vbext_ComponentType
MbrAy = ZMbrAy(Ay, NmPatn)
End Property

Property Get MbrNy(Optional NmPatn$ = ".") As String()
MbrNy = Oy(MdAy(NmPatn)).Ny
End Property

Property Get MdAy(Optional NmPatn$ = ".", Optional CmpTyAy0) As CodeModule()
Dim Ay() As vbext_ComponentType
Ay = DftCmpTyAy(CmpTyAy0)
MdAy = ZMbrAy(Ay, NmPatn)
End Property

Property Get MdAyOfStd(Optional NmPatn$ = ".") As CodeModule()
MdAyOfStd = MdAy(NmPatn, CmpTyAyOfStd)
End Property

Property Get MdNy(Optional NmPatn$ = ".", Optional CmpTyAy0) As String()
MdNy = Oy(MdAy(NmPatn, CmpTyAy0)).Ny
End Property

Property Get MdNyOfStd(Optional NmPatn$ = ".") As String()
MdNyOfStd = MdNy(NmPatn, CmpTyAyOfStd)
End Property

Property Get MthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
'Dim Fny$()
'    Fny = FnyOfMthDrs(WithBdyLy, WithBdyLines)
'MthDrs.Fny = Fny
'MthDrs.Dry = MthDry(WithBdyLy, WithBdyLines)
Stop '
End Property

Property Get MthDry(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In MdAy
        Set Md = I
        PushAy Dry, MdMthDry(Md, WithBdyLy, WithBdyLines)
    Next
MthDry = Dry
End Property

Property Get MthLinDry() As Variant()
Stop '
'Dim I, Md As CodeModule, O()
'For Each I In MbrAy
'    Set Md = I
'    Dim N$: N = MdNm(Md)
'    Dim Ay$(): Ay = MdMthLinAy(Md)
'    Dim Dry(): Dry = ConstAy_ConstValDry(N, Ay)
'    PushAy O, Dry
'Next
'MthLinDry = O
End Property

Sub ZZ_MthNy()
Stop
'AyBrw Pjx(CurPj).MthNy
End Sub
Property Get MthNy(Optional CmpTyAy0, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = ".") As String()
Dim CmpTyAy() As vbext_ComponentType
    CmpTyAy = DftCmpTyAy(CmpTyAy0)
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(MdNy(MdNmPatn))
If AyIsEmp(Ny) Then Exit Property
For Each I In Ny
    Set M = Md(CStr(I))
    Stop
'    PushAy O, AyAddPfx(MdMthNy(M, MthNmPatn), MdNm(M) & Sep)
Next
MthNy = O
End Property

Property Get PatnLy(Patn$) As String()
Dim I, Md As CodeModule, O$()
For Each I In MdAy
   Set Md = I
   Stop '
'   PushAy O, MdPatnLy(Md, Patn)
Next
PatnLy = O
End Property

Property Get Pth$()
Stop '
'Pth = FfnPth(A.Filename)
End Property

Property Get ReadRfCfg() As String()
Const CSub$ = "PjReadRfCfg"
Dim B$: B = RfCfgFfn
Stop '
'If Not FfnIsExist(B) Then Er CSub, "{Pj-Rf-Cfg-Fil} not found", B
'ReadRfCfg = FtLy(B)
End Property

Sub RenMdByPfx(FmMdPfx$, ToMdPfx$)
Dim DftNy$()
Dim Ny$()
    Ny = MdNy("^" & FmMdPfx)
    DftNy = AyMapAsgSy(Ny, "RplPfx", FmMdPfx, ToMdPfx)
Dim MdAy() As CodeModule
    Dim MdNm
    Dim M As CodeModule
    For Each MdNm In Ny
        Set M = Md(CStr(MdNm))
        PushObj MdAy, M
    Next
Dim I%, U%
    For I = 0 To UB(DftNy)
        Stop '
        'MdRen MdAy(I), DftNy(I)
    Next
End Sub

Property Get RfAy() As VBIDE.Reference()
Dim RF As VBIDE.Reference, O() As VBIDE.Reference
For Each RF In A.References
    Push O, RF
Next
RfAy = O
End Property

Sub RfBrw()
AyBrw RfLy
End Sub

Property Get RfCfgFfn$()
RfCfgFfn = SrcPth & "PjRf.Cfg"
End Property

Sub RfDmp()
AyDmp RfLy
End Sub

Property Get RfLy() As String()
Dim O$()
Dim R() As VBIDE.Reference
R = RfAy
Dim Ny$(): Ny = Oy(R).Ny
Ny = AyAlignL(Ny)
Dim J%
For J = 0 To UB(Ny)
    Push O, Ny(J) & " " & RfPth(R(J))
Next
RfLy = O
End Property

Sub RmvMdNmPfx(Pfx$)
Dim I, Md As CodeModule
For Each I In MdAy("^" & Pfx)
    Set Md = I
    Stop '
    'MdRmvNmPfx Md, Pfx
Next
End Sub

Property Get SrcPth$()
Dim Fn$: Fn = File(Ffn).Fn
Dim O$:
Stop '
'O = FfnPth(A.Filename) & "Src\": PthEns O
O = O & Fn & "\":                'PthEns O
SrcPth = O
End Property

Sub SrcPthBrw()
Stop '
'PthBrw SrcPth
End Sub

Property Get TyNy(Optional TyNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = vbTab) As String()
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(MdNy(MdNmPatn))
If AyIsEmp(Ny) Then Exit Property
For Each I In Ny
    Set M = Md(CStr(I))
    Stop '
'    PushAy O, AyAddPfx(MdTyNy(M, TyNmPatn), MdNm(M) & Sep)
Next
TyNy = O
End Property

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

Property Get EnsMd(MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
If Not HasMd(MdNm) Then
    Set EnsMd = CrtMd(MdNm, Ty)
Else
    Set EnsMd = Md(MdNm)
End If
End Property

Property Get SrcDrs() As Drs
'PjSrcDry is 2 col dry with columns MdNm and SrcLin-Class
Dim O(), L%, I, N$, M As Mdx, S
For Each I In MdxAy
    Set M = I
    N = M.Nm
    L = 0
    Stop '
'    For Each S In M.Md .Src
'        Push O, Array(N, L, SrcLin(S))
'        L = L + 1
'    Next
Next
SrcDrs = Drs("Md Lx Src", O)
End Property

Sub GoMdNm(MdNm$, Optional ClsOth As Boolean)
If ClsOth Then WinClsCd
Md(MdNm).CodePane.Show
End Sub

Property Get Md(MdNm) As CodeModule
Set Md = A.VBComponents(MdNm).CodeModule
End Property


Function DftPj(A As VBProject) As VBProject
If IsNothing(A) Then
   Set DftPj = CurPj
Else
   Set DftPj = A
End If
End Function

Function DftPjByPjNm(A$) As VBProject
If A = "" Then
   Set DftPjByPjNm = CurPjx
   Exit Function
End If
Dim I As VBProject
For Each I In CurVbe.VBProjects
   If UCase(I.Name) = UCase(A) Then Set DftPjByPjNm = I: Exit Function
Next
Stop
End Function

Function Pj(PjNm$) As VBProject
Set Pj = CurVbe.VBProjects(PjNm)
End Function

Function PjAddMd(A As VBProject, MdNm$) As CodeModule
Dim O As VBComponent
Set O = A.VBComponents.Add(vbext_ct_StdModule)
O.Name = MdNm
Set PjAddMd = O.CodeModule
End Function

Sub PjAddRf(A As VBProject, RfNm, PjFfn)
If PjHasRf(A, RfNm) Then Exit Sub
A.References.AddFromFile PjFfn
End Sub

Function PjClsAy(A As VBProject, Optional NmPatn$ = ".") As CodeModule()
PjClsAy = PjMdAy(A, NmPatn, CmpTyAyOfCls)
End Function

Function PjClsNy(A As VBProject, Optional NmPatn$ = ".") As String()
PjClsNy = OyPrp(PjClsAy(A, NmPatn), "Name", EmpSy)
End Function

Sub PjCpyToSrc(A As VBProject)
FfnCpyToPth A.Filename, PjSrcPth(A), OvrWrt:=True
End Sub

Function PjDltMd(A As VBProject, MdNm$) As Boolean
If Not PjHasMd(A, MdNm) Then Exit Function
A.VBComponents.Remove A.VBComponents(MdNm)
PjDltMd = True
End Function

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

Function PjFfn$(A As VBProject)
On Error Resume Next
PjFfn = A.Filename
End Function

Function PjHasMd(A As VBProject, MdNm$) As Boolean
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If MdNm = Cmp.Name Then
        PjHasMd = True
        Exit Function
    End If
Next
End Function

Function PjHasMdNm(A As VBProject, MdNm$) As Boolean
Dim I
Dim Cmp As VBComponent
For Each I In A.VBComponents
    Set Cmp = I
    If Cmp.Name = MdNm Then PjHasMdNm = True: Exit Function
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

Function PjMbrAy(A As VBProject, Optional NmPatn$ = ".") As CodeModule()
Dim Ay() As vbext_ComponentType
PjMbrAy = ZPjMbrAy(A, Ay, NmPatn)
End Function

Function PjMbrNy(A As VBProject, Optional NmPatn$ = ".") As String()
PjMbrNy = OyPrp(PjMdAy(A, NmPatn), "Name", EmpSy)
End Function

Function PjMdAy(A As VBProject, Optional NmPatn$ = ".", Optional CmpTyAy0) As CodeModule()
Dim Ay() As vbext_ComponentType
Ay = DftCmpTyAy(CmpTyAy0)
PjMdAy = ZPjMbrAy(A, Ay, NmPatn)
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

Sub ZZ_PjMthNy()
AyBrw PjMthNy(CurPjx)
End Sub
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
AyBrw PjRfLy(A)
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

Private Function ZPjMbrAy(A As VBProject, MbrTyAy() As vbext_ComponentType, Optional NmPatn$ = ".") As CodeModule()
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

Private Sub CurPjx__Tst()
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

