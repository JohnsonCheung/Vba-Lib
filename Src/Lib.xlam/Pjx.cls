VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pjx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A As VBProject
Property Get Init(Pj) As Pjx
Set A = Pj
Set Init = Me
End Property
Property Get Nm$()
Nm = A.Name
End Property
Sub Sort()
Dim M As CodeModule, I
For Each I In MbrAy
    Set M = I
    MdSrt M
Next
End Sub

Property Get Pj() As VBProject
Set Pj = A
End Property

Property Get ToStr$()
ToStr = Nm
End Property
Property Get IsFxa() As Boolean
IsFxa = LCase(FfnExt(Me.Ffn)) = ".xlam"
End Property
Property Get IsUsrLib() As Boolean
IsUsrLib = True
If IsFxa Then Exit Property
IsUsrLib = True
End Property
Property Get AddMd(MdNm$) As CodeModule
Dim O As VBComponent
Set O = A.VBComponents.Add(vbext_ct_StdModule)
O.Name = MdNm
Set AddMd = O.CodeModule
End Property

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
FfnCpyToPth A.Filename, SrcPth, OvrWrt:=True
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
FfnCpyToPth A.Filename, SrcPth, OvrWrt:=True
PthClrFil SrcPth 'Clr SrcPth ---
'Oy(MbrAy).EachSub "Exp"
Dim I
For Each I In MbrAy
    MdExp CvMd(I)  'Exp each md --
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
   MdRmvOptCmpDb CvMd(I)
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
Dim RF As VBIde.Reference
For Each RF In A.References
    If RF.Name = RfNm Then HasRf = True: Exit Property
Next
End Property

Sub ImpRf(RfCfgPth$)
Dim B As Dictionary: Set B = FtDic(RfCfgPth & "PjRf.Cfg")
Dim K
For Each K In B.Keys
    AddRf K, B(K)
Next
End Sub

Property Get IsUnderSrcPth() As Boolean
If PthFdr(Pth) = "Src" Then Stop
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
Dim Fny$()
    Fny = FnyOfMthDrs(WithBdyLy, WithBdyLines)
MthDrs.Fny = Fny
MthDrs.Dry = MthDry(WithBdyLy, WithBdyLines)
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
Dim I, Md As CodeModule, O()
For Each I In MbrAy
    Set Md = I
    Dim N$: N = MdNm(Md)
    Dim Ay$(): Ay = MdMthLinAy(Md)
    Dim Dry(): Dry = ConstAy_ConstValDry(N, Ay)
    PushAy O, Dry
Next
MthLinDry = O
End Property

Sub ZZ_MthNy()
AyBrw Pjx(CurPj).MthNy
End Sub
Property Get MthNy(Optional CmpTyAy0, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = ".") As String()
Dim CmpTyAy() As vbext_ComponentType
    CmpTyAy = DftCmpTyAy(CmpTyAy0)
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(MdNy(MdNmPatn))
If AyIsEmp(Ny) Then Exit Property
For Each I In Ny
    Set M = Md(CStr(I))
    PushAy O, AyAddPfx(MdMthNy(M, MthNmPatn), MdNm(M) & Sep)
Next
MthNy = O
End Property

Property Get PatnLy(Patn$) As String()
Dim I, Md As CodeModule, O$()
For Each I In MdAy
   Set Md = I
   PushAy O, MdPatnLy(Md, Patn)
Next
PatnLy = O
End Property

Property Get Pth$()
Pth = FfnPth(A.Filename)
End Property

Property Get ReadRfCfg() As String()
Const CSub$ = "PjReadRfCfg"
Dim B$: B = RfCfgFfn
If Not FfnIsExist(B) Then Er CSub, "{Pj-Rf-Cfg-Fil} not found", B
ReadRfCfg = FtLy(B)
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
        MdRen MdAy(I), DftNy(I)
    Next
End Sub

Property Get RfAy() As VBIde.Reference()
Dim RF As VBIde.Reference, O() As VBIde.Reference
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
Dim R() As VBIde.Reference
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
    MdRmvNmPfx Md, Pfx
Next
End Sub

Property Get SrcPth$()
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(A.Filename) & "Src\": PthEns O
O = O & Fn & "\":                PthEns O
SrcPth = O
End Property

Sub SrcPthBrw()
PthBrw SrcPth
End Sub

Property Get TyNy(Optional TyNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = vbTab) As String()
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(MdNy(MdNmPatn))
If AyIsEmp(Ny) Then Exit Property
For Each I In Ny
    Set M = Md(CStr(I))
    PushAy O, AyAddPfx(MdTyNy(M, TyNmPatn), MdNm(M) & Sep)
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
Dim O(), L%, I, N$, M As CodeModule, S
For Each I In MdAy
    Set M = I
    N = MdNm(M)
    L = 0
    For Each S In MdSrc(M)
        Push O, Array(N, L, SrcLin(S))
        L = L + 1
    Next
Next
SrcDrs = NewDrs("Md Lx Src", O)
End Property

Sub GoMdNm(MdNm$, Optional ClsOth As Boolean)
If ClsOth Then WinClsCd
Md(MdNm).CodePane.Show
End Sub

Property Get Md(MdNm) As CodeModule
Set Md = A.VBComponents(MdNm).CodeModule
End Property

