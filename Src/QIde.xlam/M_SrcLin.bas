Attribute VB_Name = "M_SrcLin"
Option Explicit
Private A$
Private A_IsMth As Boolean

Function EnmNm$()
If IsEmn Then EnmNm = Lin(NoEnm).Nm
End Function

Function FriendMthLin$()
If IsMth Then FriendMthLin = "Friend " & NoMdy
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

'Property Get MthBrk() As MthBrk
'If Not IsMth Then Exit Property
'Dim O As MthBrk
'With O
'    .Mdy = Mdy
'    .MthNm = MthNm
'    .Ty = MthTy
'End With
'MthBrk = O
'End Property

Property Get MthNm$()
If IsMth Then MthNm = Lin(NoMthTy).Nm
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
If IsPrpLin Then PrpTy = Lin(NoFunTy).T1
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
If IsTy Then TyNm = Lin(NoTy).Nm
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
Dim MNm$, X As SrcLin
For Each I In CurPjx.MbrAy
    Set Md = I
    MNm = MdNm(Md)
    Lno = 0
    For Each Lin In MdSrc(Md)
        Lno = Lno + 1
        Set X = Ide.SrcLin(Lin)
        Push Dry, X.InfDr(MNm, Lno)
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
'Init ZZMthLin
AllSrcCode__Tst
'MthBrk__Tst
'IsMth__Tst
End Sub


Public Const C_Enm$ = "Enum"
Public Const C_Prp$ = "Property"
Public Const C_Ty$ = "Type"
Public Const C_Fun$ = "Function"
Public Const C_Sub$ = "Sub"
Public Const C_Get$ = "Get"
Public Const C_Set$ = "Set"
Public Const C_Let$ = "Let"
Public Const C_Pub$ = "Public"
Public Const C_Prv$ = "Private"
Public Const C_Frd$ = "Friend"
Public Const C_PrpGet$ = C_Prp + " " + C_Get
Public Const C_PrpLet$ = C_Prp + " " + C_Let
Public Const C_PrpSet$ = C_Prp + " " + C_Set

Function KwIsFunTy(S) As Boolean
KwIsFunTy = AyHas(SyOfFunTy, S)
End Function

Function KwIsMdy(Mdy) As Boolean
KwIsMdy = AyHas(Array("Private", "Public", "Friend", ""), Mdy)
End Function

Function KwIsMthTy(S) As Boolean
KwIsMthTy = AyHas(S, SyOfMthTy)
End Function

Function MthLin_EnsPrivate(A) As SomStr
Dim P As Parse: P = ParseKwMdy(NewParse(A))
If Not P.IsOk Then Exit Function
Dim P1 As Parse: P1 = ParseKwMthTy(P)
If Not P.IsOk Then Exit Function
If P.Er_or_Ok = "Private" Then MthLin_EnsPrivate = SomStr(A): Exit Function
MthLin_EnsPrivate = SomStr("Private " & P.Lin)
End Function

Function MthLin_Key$(A)
With SrcLin_MthBrk(A)
    MthLin_Key = FmtQQ("?:?:?", .Mdy, .Ty, .MthNm)
End With
End Function
Function SrcLin_EndLinPfx$(A)
Ass SrcLin_IsMth(A)
Stop '
'SrcLin_EndLinPfx = "End " & Lin(SrcLin_MthTy(A)).T1
End Function

Function SrcLin_EnmNm$(A)
'If SrcLin_IsEmn(A) Then EnmNm = LinNm(NoEnm)
End Function

Function SrcLin_IsCd(A) As Boolean
If Linx(A).IsEmp Then Exit Function
If SrcLin_IsRmk(A) Then Exit Function
SrcLin_IsCd = True
End Function

Function SrcLin_IsEmn(Lin) As Boolean
SrcLin_IsEmn = HasPfx(SrcLin_RmvMdy(Lin), "Enum")
End Function

Function SrcLin_IsMth(A) As Boolean
Stop '
'If HasPfx(A, "Function") Then Stop
'SrcLin_IsMth = KwIsFunTy(Lin(SrcLin_RmvMdy(A)).T1)
End Function

Function SrcLin_IsRmk(Lin) As Boolean
SrcLin_IsRmk = FstChr(LTrim(Lin)) = "'"
End Function

Function SrcLin_IsTy(Lin) As Boolean
SrcLin_IsTy = HasPfx(SrcLin_RmvMdy(Lin), "Type")
End Function

Function SrcLin_Mdy$(A)
SrcLin_Mdy = ParseRet(ParseKwMdy(NewParse(A)))
End Function

Function SrcLin_MthBrk(A) As MthBrk
Dim P As Parse
P = ParseKwMdy(NewParse(A)): If P.IsOk Then SrcLin_MthBrk.Mdy = P.Er_or_Ok Else Exit Function
P = ParseKwMthTy(P):         If P.IsOk Then SrcLin_MthBrk.Ty = P.Er_or_Ok Else Exit Function
P = ParseNm(P):              If P.IsOk Then SrcLin_MthBrk.MthNm = P.Er_or_Ok
End Function

Function SrcLin_MthDr(A, Lno&, Optional MdNm$, Optional MdTy$) As Variant()
With SrcLin_MthBrk(A)
   SrcLin_MthDr = Array(MdNm, Lno, .Mdy, .Ty, .MthNm)
End With
End Function

Function SrcLin_MthNm$(A)
SrcLin_MthNm = ParseRet(ParseNm(ParseKwMthTy(ParseKwMdy(NewParse(A)))))
End Function

Function SrcLin_MthTy$(A)
SrcLin_MthTy = SrcLin_MthBrk(A).Ty
End Function

Function SrcLin_RmvMdy$(A)
SrcLin_RmvMdy = LTrim(RmvPfxAy(A, SyOfMdy))
End Function

Function SrcLin_TyNm$(A)
SrcLin_TyNm = ParseRet(ParseNm(ParseKwTy(ParseKwMdy(NewParse(A)))))
End Function

Function SyOfFunTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Fun, C_Sub, C_Prp)
End If
SyOfFunTy = Y
End Function

Function SyOfMdy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Pub, C_Prv, C_Frd)
End If
SyOfMdy = Y
End Function

Function SyOfMthTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Fun, C_Sub, C_PrpGet, C_PrpLet, C_PrpSet)
End If
SyOfMthTy = Y
End Function

Function SyOfPrpTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = ApSy(C_Get, C_Set, C_Let)
End If
SyOfPrpTy = Y
End Function

Function SyOfSrcTy() As String()
Static X As Boolean, Y
If Not X Then
   X = True
   Y = SyOfMthTy
   Push Y, C_Ty
   Push Y, C_Enm
End If
SyOfSrcTy = Y
End Function

Private Sub ZZ_PrjSrcDrs()
Dim O As Drs: O = CurPjx.SrcDrs
'DryBrw O

Dim A As SrcLin: Set A = V(O.Dry(2)(1)).SrcLin
Dim A1 As Drs: A1 = A.InfDrs
DrsDmp A1
Stop
End Sub

Private Function ZZSrc() As String()
'ZZSrc = MdSrc(Md("IdeSrcLin"))
End Function

Private Function ZZSrcLin$()
ZZSrcLin = "Private Sub SrcLin_IsMth()"
End Function

Private Sub ZZ_SrcLin_IsMth()
Dim O()
Dim L
For Each L In ZZSrc
    Push O, Array(IIf(SrcLin_IsMth(L), "*Mth", ""), MthLin_Key(L), L)
Next
DrsBrw NewDrs("IsMth Key Lin", O)
End Sub

Private Sub SrcLin_IsMth__Tst()
ZZ_SrcLin_IsMth
End Sub

Private Sub SrcLin_MthBrk__Tst()
Dim Act As MthBrk:
Act = SrcLin_MthBrk("Private Function AA()")
Ass Act.Mdy = "Private"
Ass Act.Ty = "Function"
Ass Act.MthNm = "AA"

Act = SrcLin_MthBrk("Private Sub TakBet__Tst()")
Ass Act.Mdy = "Private"
Ass Act.Ty = "Sub"
Ass Act.MthNm = "TakBet__Tst"
End Sub

Private Sub SrcLin_MthNm__Tst()
Dim Act$
Dim Lin$
Lin = "Private Sub SrcLin_MthNm__Tst )": Act = SrcLin_MthNm(Lin): Ass Act = "SrcLin_MthNm__Tst"
Lin = "Property Set A(V)":           Act = SrcLin_MthNm(Lin): Ass Act = "A"
End Sub

