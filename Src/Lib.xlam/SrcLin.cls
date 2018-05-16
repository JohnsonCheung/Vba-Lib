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
Private A_SrcLin_IsMth As Boolean
Property Get EnmNm$()
If SrcLin_IsEmn Then EnmNm = LinNm(NoEnm)
End Property

Function FriendMthLin$()
If SrcLin_IsMth Then FriendMthLin = "Friend " & NoMdy
End Function

Function Init(Lin) As SrcLin
A = Lin
A_SrcLin_IsMth = HasOneOfPfx(NoMdy, SyOfMthTy)
Set Init = Me
End Function

Property Get MthBrk() As MthBrk
If Not SrcLin_IsMth Then Exit Property
Dim O As MthBrk
With O
    .Mdy = Mdy
    .MthNm = MthNm
    .Ty = MthTy
End With
MthBrk = O
End Property

Property Get LinIsEmp() As Boolean
LinIsEmp = Trim(A) = ""
End Property

Property Get SrcLin_IsEmn() As Boolean
SrcLin_IsEmn = HasPfx(NoMdy, C_Enm)
End Property
Sub MthBrk__Tst()
With MthBrk
    Debug.Print .Mdy
    Debug.Print .MthNm
    Debug.Print .Ty
End With
End Sub
Private Function ZZMthLin$()
ZZMthLin = "Property Get AA()"
End Function
Property Get SrcLin_IsMth() As Boolean
SrcLin_IsMth = A_SrcLin_IsMth
End Property

Property Get IsPrpLin() As Boolean
IsPrpLin = HasPfx(NoMdy, C_Prp)
End Property

Property Get SrcLin_IsRmk() As Boolean
SrcLin_IsRmk = FstChr(LTrim(A)) = "'"
End Property

Property Get SrcLin_IsTy() As Boolean
SrcLin_IsTy = HasPfx(NoMdy, C_Ty)
End Property

Property Get Mdy$()
Mdy = StrPfx(A, SyOfMdy)
End Property

Property Get MthNm$()
If SrcLin_IsMth Then MthNm = LinNm(NoMthTy)
End Property

Property Get MthTy$()
If SrcLin_IsMth Then MthTy = StrPfx(NoMdy, SyOfMthTy)
End Property

Function NoMdy$()
NoMdy = LTrim(RmvPfxAy(A, SyOfMdy))
End Function

Function PrivateMthLin$()
If SrcLin_IsMth Then PrivateMthLin = "Private " & NoMdy
End Function

Property Get PrpTy$()
If IsPrpLin Then PrpTy = LinT1(NoFunTy)
End Property

Function PrpValDr() As Variant()
PrpValDr = Array(, , A, EnmNm, LinIsEmp, SrcLin_IsEmn, SrcLin_IsMth, IsPrpLin, SrcLin_IsRmk, SrcLin_IsTy, Mdy, MthNm, MthTy, NoMdy, PrpTy, TyNm)
End Function

Property Get PrpValFny() As String()
Static X As Boolean, Y$()
If Not X Then
    X = True
    Y = LvsSy("Md Lno Lin EnmNm LinIsEmp SrcLin_IsEmn SrcLin_IsMth IsPrpLin SrcLin_IsRmk SrcLin_IsTy Mdy MthNm MthTy NoMdy PrpTy TyNm")
End If
PrpValFny = Y
End Property

Function PublicMthLin$()
If SrcLin_IsMth Then PublicMthLin = NoMdy
End Function

Property Get TyNm$()
If SrcLin_IsTy Then TyNm = LinNm(NoTy)
End Property

Private Property Get NoEnm$()
If SrcLin_IsEmn Then NoEnm = LTrim(RmvPfx(NoMdy, C_Enm))
End Property

Private Property Get NoFunTy$()
If SrcLin_IsMth Then NoFunTy = RmvPfxAy(NoMdy, SyOfFunTy)
End Property

Private Property Get NoMthTy$()
If SrcLin_IsMth Then NoMthTy = LTrim(RmvPfxAy(NoMdy, SyOfMthTy))
End Property

Private Property Get NoTy$()
If SrcLin_IsTy Then NoTy = LTrim(RmvPfx(NoMdy, C_Ty))
End Property
Private Sub AllSrcCode__Tst()
Dim Dry()
Dim Dr()
Dim Drs As Drs
Dim O$()
Dim I, Lin
Dim Md As CodeModule:
Dim Lno&
Dim MNm$
For Each I In PjMbrAy(CurPj)
    Set Md = I
    MNm = MdNm(Md)
    Lno = 0
    For Each Lin In MdSrc(Md)
        Lno = Lno + 1
        A = Lin
        Dr = PrpValDr
        Dr(0) = MNm
        Dr(1) = Lno
        Push Dry, Dr
    Next
Next
Drs.Dry = Dry
Drs.Fny = PrpValFny
DrsWs Drs
End Sub

Private Sub SrcLin_IsMth__Tst()
A = ZZMthLin
Ass SrcLin_IsMth = True
End Sub

Sub Tst()
Init ZZMthLin
AllSrcCode__Tst
MthBrk__Tst
SrcLin_IsMth__Tst
End Sub
