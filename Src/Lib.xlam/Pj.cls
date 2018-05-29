VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Pj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A As VBProject
Function Init(Pj As VBProject) As Pj
Set A = Pj
Set Init = Me
End Function
Function Nm$()
Nm = A.Name
End Function
Function Pj() As VBProject
Set Pj = A
End Function
Function Str$()
Str = Nm
End Function
Function AddMd(MdNm$) As CodeModule
Dim O As VBComponent
Set O = A.VBComponents.Add(vbext_ct_StdModule)
O.Name = MdNm
Set AddMd = O.CodeModule
End Function

Sub AddRf(RfNm, PjFfn)
If HasRf(RfNm) Then Exit Sub
A.References.AddFromFile PjFfn
End Sub

Function ClsAy(Optional NmPatn$ = ".") As CodeModule()
ClsAy = MdAy(NmPatn, CmpTyAyOfCls)
End Function

Function ClsNy(Optional NmPatn$ = ".") As String()
ClsNy = OyNy(ClsAy(NmPatn))
End Function

Sub CpyToSrc()
FfnCpyToPth A.Filename, SrcPth, OvrWrt:=True
End Sub

Function CrtMd(Optional MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
If MdNm <> "" Then
    If HasMd(MdNm) Then
        Er "NewMd", "Given {MdNm} exist", MdNm
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

Function Ffn$()
On Error Resume Next
Ffn = A.Filename
End Function

Sub ImpSrcFfn(SrcFfn)
A.VBComponents.Import SrcFfn
End Sub

Sub RmvOptCmpDbLin()
Dim I
For Each I In MdAy
   MdRmvOptCmpDb CvMd(I)
Next
End Sub
Function HasMd(MdNm$) As Boolean
Dim Cmp As VBComponent
For Each Cmp In A.VBComponents
    If MdNm = Cmp.Name Then
        HasMd = True
        Exit Function
    End If
Next
End Function

Function HasMdNm(MdNm$) As Boolean
Dim I
Dim Cmp As VBComponent
For Each I In A.VBComponents
    Set Cmp = I
    If Cmp.Name = MdNm Then HasMdNm = True: Exit Function
Next
End Function

Function HasRf(RfNm)
Dim RF As VBIDE.Reference
For Each RF In A.References
    If RF.Name = RfNm Then HasRf = True: Exit Function
Next
End Function

Sub ImpRf(RfCfgPth$)
Dim B As Dictionary: Set B = FtDic(RfCfgPth & "PjRf.Cfg")
Dim K
For Each K In B.Keys
    AddRf K, B(K)
Next
End Sub

Function IsUnderSrcPth() As Boolean
If PthFdr(Pth) = "Src" Then Stop
End Function

Function MbrAy(Optional NmPatn$ = ".") As CodeModule()
Dim Ay() As vbext_ComponentType
MbrAy = ZMbrAy(Ay, NmPatn)
End Function

Function MbrNy(Optional NmPatn$ = ".") As String()
MbrNy = OyNy(MdAy(NmPatn))
End Function

Function MdAy(Optional NmPatn$ = ".", Optional CmpTyAy0) As CodeModule()
Dim Ay() As vbext_ComponentType
Ay = DftCmpTyAy(CmpTyAy0)
MdAy = ZMbrAy(Ay, NmPatn)
End Function

Function MdAyOfStd(Optional NmPatn$ = ".") As CodeModule()
MdAyOfStd = MdAy(NmPatn, CmpTyAyOfStd)
End Function

Function MdNy(Optional NmPatn$ = ".", Optional CmpTyAy0) As String()
MdNy = OyNy(MdAy(NmPatn, CmpTyAy0))
End Function

Function MdNyOfStd(Optional NmPatn$ = ".") As String()
MdNyOfStd = MdNy(NmPatn, CmpTyAyOfStd)
End Function

Function MthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Drs
Dim Fny$()
    Fny = FnyOfMthDrs(WithBdyLy, WithBdyLines)
MthDrs.Fny = Fny
MthDrs.Dry = MthDry(WithBdyLy, WithBdyLines)
End Function

Function MthDry(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As Variant()
Dim Dry()
    Dim I, Md As CodeModule
    For Each I In MdAy
        Set Md = I
        PushAy Dry, MdMthDry(Md, WithBdyLy, WithBdyLines)
    Next
MthDry = Dry
End Function

Function MthLinDry() As Variant()
Dim I, Md As CodeModule, O()
For Each I In MbrAy
    Set Md = I
    Dim N$: N = MdNm(Md)
    Dim Ay$(): Ay = MdMthLinAy(Md)
    Dim Dry(): Dry = ConstAy_ConstValDry(N, Ay)
    PushAy O, Dry
Next
MthLinDry = O
End Function

Sub ZZ_MthNy()
AyBrw Pj(CurPj).MthNy
End Sub
Function MthNy(Optional CmpTyAy0, Optional MthNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = ".") As String()
Dim CmpTyAy() As vbext_ComponentType
    CmpTyAy = DftCmpTyAy(CmpTyAy0)
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(MdNy(MdNmPatn))
If AyIsEmp(Ny) Then Exit Function
For Each I In Ny
    Set M = Md(CStr(I))
    PushAy O, AyAddPfx(MdMthNy(M, MthNmPatn), MdNm(M) & Sep)
Next
MthNy = O
End Function

Function PatnLy(Patn$) As String()
Dim I, Md As CodeModule, O$()
For Each I In MdAy
   Set Md = I
   PushAy O, MdPatnLy(Md, Patn)
Next
PatnLy = O
End Function

Function Pth$()
Pth = FfnPth(A.Filename)
End Function

Function ReadRfCfg() As String()
Const CSub$ = "PjReadRfCfg"
Dim B$: B = RfCfgFfn
If Not FfnIsExist(B) Then Er CSub, "{Pj-Rf-Cfg-Fil} not found", B
ReadRfCfg = FtLy(B)
End Function

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

Function RfAy() As VBIDE.Reference()
Dim RF As VBIDE.Reference, O() As VBIDE.Reference
For Each RF In A.References
    Push O, RF
Next
RfAy = O
End Function

Sub RfBrw()
AyBrw RfLy
End Sub

Function RfCfgFfn$()
RfCfgFfn = SrcPth & "PjRf.Cfg"
End Function

Sub RfDmp()
AyDmp RfLy
End Sub

Function RfLy() As String()
Dim O$()
Dim R() As VBIDE.Reference
R = RfAy
Dim Ny$(): Ny = OyNy(R)
Ny = AyAlignL(Ny)
Dim J%
For J = 0 To UB(Ny)
    Push O, Ny(J) & " " & RfPth(R(J))
Next
RfLy = O
End Function

Sub RmvMdNmPfx(Pfx$)
Dim I, Md As CodeModule
For Each I In MdAy("^" & Pfx)
    Set Md = I
    MdRmvNmPfx Md, Pfx
Next
End Sub

Function SrcPth$()
Dim Fn$: Fn = FfnFn(Ffn)
Dim O$:
O = FfnPth(A.Filename) & "Src\": PthEns O
O = O & Fn & "\":                PthEns O
SrcPth = O
End Function

Sub SrcPthBrw()
PthBrw SrcPth
End Sub

Function TyNy(Optional TyNmPatn$ = ".", Optional MdNmPatn$ = ".", Optional Sep$ = vbTab) As String()
Dim O$(), I, M As CodeModule, Ay$(), Ny$()
Ny = AySrt(MdNy(MdNmPatn))
If AyIsEmp(Ny) Then Exit Function
For Each I In Ny
    Set M = Md(CStr(I))
    PushAy O, AyAddPfx(MdTyNy(M, TyNmPatn), MdNm(M) & Sep)
Next
TyNy = O
End Function

Private Function CmpTyAyOfCls() As vbext_ComponentType()
Dim T(0) As vbext_ComponentType
T(0) = vbext_ct_ClassModule
CmpTyAyOfCls = T
End Function

Private Function CmpTyAyOfStd() As vbext_ComponentType()
Dim Ay(0) As vbext_ComponentType
Ay(0) = vbext_ct_StdModule
CmpTyAyOfStd = Ay
End Function

Private Function ZMbrAy(MbrTyAy() As vbext_ComponentType, Optional NmPatn$ = ".") As CodeModule()
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
ZMbrAy = O
End Function

Function EnsMd(MdNm$, Optional Ty As vbext_ComponentType = vbext_ct_StdModule) As CodeModule
If Not HasMd(MdNm) Then
    Set EnsMd = CrtMd(MdNm, Ty)
Else
    Set EnsMd = Md(MdNm)
End If
End Function

Function SrcDrs() As Drs
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
End Function

Sub GoMdNm(MdNm$, Optional ClsOth As Boolean)
If ClsOth Then WinClsCd
Md(MdNm).CodePane.Show
End Sub

Function Md(MdNm) As CodeModule
Set Md = A.VBComponents(MdNm).CodeModule
End Function

