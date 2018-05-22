VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SrcLin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A$
Private A_IsMth As Boolean

Function EnmNm$()
If IsEmn Then EnmNm = LinNm(NoEnm)
End Function

Function FriendMthLin$()
If IsMth Then FriendMthLin = "Friend " & NoMdy
End Function

Function Init(Lin) As SrcLin
A = Lin
Dim NoMdy$: NoMdy = LTrim(RmvPfxAy(A, SyOfMdy))
A_IsMth = HasOneOfPfx(NoMdy, SyOfMthTy)
Set Init = Me
End Function

Property Get IsPrpLin() As Boolean
IsPrpLin = HasPfx(NoMdy, C_Prp)
End Property

Property Get IsBlank() As Boolean
IsBlank = Trim(A) = ""
End Property

Property Get Mdy$()
Mdy = StrPfx(A, SyOfMdy)
End Property

Property Get MthBrk() As MthBrk
If Not IsMth Then Exit Property
Dim O As MthBrk
With O
    .Mdy = Mdy
    .MthNm = MthNm
    .Ty = MthTy
End With
MthBrk = O
End Property

Property Get MthNm$()
If IsMth Then MthNm = LinNm(NoMthTy)
End Property

Property Get MthTy$()
If IsMth Then MthTy = StrPfx(NoMdy, SyOfMthTy)
End Property

Property Get NoMdy$()
If Not A_IsMth Then Exit Property
NoMdy = LTrim(RmvPfxAy(A, SyOfMdy))
End Property

Property Get PrivateMthLin$()
If IsMth Then PrivateMthLin = "Private " & NoMdy
End Property

Property Get PrpTy$()
If IsPrpLin Then PrpTy = LinT1(NoFunTy)
End Property
Function InfDrs(Optional MdNm$, Optional Lno) As Drs
With InfDrs
    .Fny = InfFny
    .Dry = Array(InfDr(MdNm, Lno))
End With
End Function
Function InfDr(MdNm$, Lno) As Variant()
InfDr = Array(MdNm, Lno, A, EnmNm, IsBlank, IsEmn, IsMth, IsPrpLin, IsRmk, IsTy, Mdy, MthNm, MthTy, NoMdy, PrpTy, TyNm)
End Function

Function InfFny() As String()
Static X As Boolean, Y$()
If Not X Then
    X = True
    Y = LvsSy("Md Lno Lin EnmNm IsBlank IsEmn IsMth IsPrpLin IsRmk IsTy Mdy MthNm MthTy NoMdy PrpTy TyNm")
End If
InfFny = Y
End Function

Function PublicMthLin$()
If IsMth Then PublicMthLin = NoMdy
End Function

Function IsEmn() As Boolean
IsEmn = HasPfx(NoMdy, C_Enm)
End Function

Function IsMth() As Boolean
IsMth = A_IsMth
End Function

Function IsRmk() As Boolean
IsRmk = FstChr(LTrim(A)) = "'"
End Function

Property Get IsTy() As Boolean
IsTy = HasPfx(NoMdy, C_Ty)
End Property
Function TyNm$()
If IsTy Then TyNm = LinNm(NoTy)
End Function

Private Function NoEnm$()
If IsEmn Then NoEnm = LTrim(RmvPfx(NoMdy, C_Enm))
End Function

Private Property Get NoFunTy$()
If IsMth Then NoFunTy = RmvPfxAy(NoMdy, SyOfFunTy)
End Property

Private Property Get NoMthTy$()
If IsMth Then NoMthTy = LTrim(RmvPfxAy(NoMdy, SyOfMthTy))
End Property

Private Property Get NoTy$()
If IsTy Then NoTy = LTrim(RmvPfx(NoMdy, C_Ty))
End Property

Private Function ZZMthLin$()
ZZMthLin = "Property Get AA()"
End Function

Private Sub AllSrcCode__Tst()
Dim Dry()
Dim Dr()
Dim Drs As Drs
Dim O$()
Dim I, Lin
Dim Md As CodeModule:
Dim Lno&
Dim MNm$
For Each I In CurPjx.MbrAy
    Set Md = I
    MNm = MdNm(Md)
    Lno = 0
    For Each Lin In MdSrc(Md)
        Lno = Lno + 1
        Init Lin
        Push Dry, InfDr(MNm, Lno)
    Next
Next
Drs.Dry = Dry
Drs.Fny = InfFny
DrsWs Drs
End Sub

Sub MthBrk__Tst()
With MthBrk
    Debug.Print .Mdy
    Debug.Print .MthNm
    Debug.Print .Ty
End With
End Sub

Private Sub IsMth__Tst()
A = ZZMthLin
Ass IsMth = True
End Sub

Sub Tst()
Init ZZMthLin
AllSrcCode__Tst
MthBrk__Tst
IsMth__Tst
End Sub

