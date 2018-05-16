Attribute VB_Name = "IdeSrcLin"
Option Explicit
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
Type Parse
    Lin As String
    IsOk As Boolean
    Er_or_Ok As String
End Type
Function NewErParse(Er$, Lin$) As Parse
NewErParse.Er_or_Ok = Er
NewErParse.Lin = Lin
End Function
Function NewOkParse(Ok$, Lin$) As Parse
Dim O As Parse
With O
    .Er_or_Ok = Ok
    .IsOk = True
    .Lin = Lin
End With
NewOkParse = O
End Function


Function SrcLin_IsRmk(Lin) As Boolean
SrcLin_IsRmk = FstChr(LTrim(Lin)) = "'"
End Function

Function NewParse(Lin) As Parse
NewParse.Lin = Lin
NewParse.IsOk = True
End Function

Function ParOneTerm(A As Parse, TermAy$()) As Parse
If Not A.IsOk Then ParOneTerm = A: Exit Function
Dim F$: F = StrPfx(A.Lin, TermAy)
If F = "" Then
   Dim Msg$: Msg = FmtQQ("These Terms[?] not found", JnVBar(TermAy))
   ParOneTerm = NewErParse(Msg, A.Lin)
Else
   ParOneTerm = NewOkParse(F, LTrim(RmvPfx(A.Lin, F)))
End If
End Function

Sub ParseBrw(A As Parse)
AyBrw ParseToLy(A)
End Sub

Sub ParseDmp(A As Parse)
AyDmp ParseToLy(A)
End Sub

Function ParseKwBktPair(A As Parse) As Parse
ParseKwBktPair = ParseStr(A, "()")
End Function

Function ParseKwEnm(A As Parse) As Parse
ParseKwEnm = ParseTerm(A, "Enum")
End Function

Function ParseKwMdy(A As Parse) As Parse
ParseKwMdy = ParseOptOneTerm(A, SyOfMdy)
End Function

Function ParseKwMthTy(A As Parse) As Parse
ParseKwMthTy = ParseOneTerm(A, SyOfMthTy)
End Function

Function ParseKwOptBktPair(A As Parse) As Parse
ParseKwOptBktPair = ParseOpt(ParseKwBktPair(A))
End Function

Function ParseKwOptional(A As Parse) As Parse
ParseKwOptional = ParseTerm(A, "Optional")
End Function

Function ParseKwPrmAy(A As Parse) As Parse
ParseKwPrmAy = ParseTerm(A, "ParamArray")
End Function

Function ParseKwTy(A As Parse) As Parse
ParseKwTy = ParseTerm(A, "Type")
End Function

Function ParseKwTyChr(A As Parse) As Parse
ParseKwTyChr = ParseOptOneChr(A, TyChrLis)
End Function

Function ParseNm(A As Parse) As Parse
If Not A.IsOk Then ParseNm = A: Exit Function
Dim B$
   B = LinNm(A.Lin)

Dim L&: L = Len(B)
If L = 0 Then
   ParseNm = NewErParse("No name", A.Lin)
Else
   ParseNm = NewOkParse(B, Mid(A.Lin, L + 1))
End If
End Function

Function ParseOneChr(A As Parse, ChrLis$) As Parse
If Not A.IsOk Then ParseOneChr = A: Exit Function
Dim C$: C = FstChr(A.Lin)
If HasSubStr(ChrLis, C) Then
   ParseOneChr = NewOkParse(C, RmvFstChr(A.Lin))
Else
   ParseOneChr = NewErParse(FmtQQ("One of ChrLis[?] not found", ChrLis), A.Lin)
End If
End Function

Function ParseOneTerm(A As Parse, TermAy$()) As Parse
Dim O As Parse
Dim J%
For J = 0 To UB(TermAy)
   O = ParseTerm(A, TermAy(J))
   If O.IsOk Then ParseOneTerm = O: Exit Function
Next
Dim Msg$
   Msg = FmtQQ("These Term[?] not found", JnSpc(TermAy))
ParseOneTerm = NewErParse(Msg, A.Lin)
End Function

Function ParseOpt(A As Parse) As Parse
If A.IsOk Then ParseOpt = A: Exit Function
A.IsOk = True
A.Er_or_Ok = ""
ParseOpt = A
End Function

Function ParseOptOneChr(A As Parse, ChrLis$) As Parse
ParseOptOneChr = ParseOpt(ParseOneChr(A, ChrLis))
End Function

Function ParseOptOneTerm(A As Parse, TermAy$()) As Parse
ParseOptOneTerm = ParseOpt(ParOneTerm(A, TermAy))
End Function

Function ParseRet$(A As Parse)
If A.IsOk Then ParseRet = A.Er_or_Ok
End Function

Function ParseRmvSpc(A As Parse) As Parse
If Not A.IsOk Then ParseRmvSpc = A: Exit Function
A.Lin = LTrim(A.Lin)
ParseRmvSpc = A
End Function

Function ParseStr(A As Parse, Str$) As Parse
If Not A.IsOk Then ParseStr = A: Exit Function
If Not HasPfx(A.Lin, Str) Then ParseStr = NewErParse(FmtQQ("[?] not found", Str), A.Lin): Exit Function
ParseStr = NewOkParse(Str, RmvPfx(A.Lin, Str))
End Function
Function SrcLin_IsEmn(Lin) As Boolean
SrcLin_IsEmn = HasPfx(SrcLin_RmvMdy(Lin), "Enum")
End Function

Function KwIsMdy(Mdy) As Boolean
KwIsMdy = AyHas(Array("Private", "Public", "Friend", ""), Mdy)
End Function

Function SrcLin_IsTy(Lin) As Boolean
SrcLin_IsTy = HasPfx(SrcLin_RmvMdy(Lin), "Type")
End Function

Private Sub SrcLin_IsMth__Tst()
ZZ_SrcLin_IsMth
End Sub
Private Sub ZZ_SrcLin_IsMth()
Dim O()
Dim L
For Each L In ZZSrc
    Push O, Array(IIf(SrcLin_IsMth(L), "*Mth", ""), MthLin_Key(L), L)
Next
DrsBrw NewDrs("IsMth Key Lin", O)
End Sub
Private Function ZZSrc() As String()
ZZSrc = MdSrc(Md("IdeSrcLin"))
End Function
Function SrcLin_IsMth(A) As Boolean
'If HasPfx(A, "Function") Then Stop
SrcLin_IsMth = KwIsFunTy(LinT1(SrcLin_RmvMdy(A)))
End Function
Function MthLin_Key$(A)
With SrcLin_MthBrk(A)
    MthLin_Key = FmtQQ("?:?:?", .Mdy, .Ty, .MthNm)
End With
End Function
Function SrcLin_RmvMdy$(A)
SrcLin_RmvMdy = LTrim(RmvPfxAy(A, SyOfMdy))
End Function
Function KwIsMthTy(S) As Boolean
KwIsMthTy = AyHas(S, SyOfMthTy)
End Function
Function KwIsFunTy(S) As Boolean
KwIsFunTy = AyHas(SyOfFunTy, S)
End Function

Function ParseTerm(A As Parse, Term$) As Parse
ParseTerm = ParseRmvSpc(ParseStr(A, Term))
End Function

Function ParseToDic(A As Parse) As Dictionary
Dim O As New Dictionary
With O
   .Add "NewParse", A.Lin
   .Add "Is", IIf(A.IsOk, "Ok", "Er")
   .Add IIf(A.IsOk, "Rslt", "Er"), A.Er_or_Ok
End With
Set ParseToDic = O
End Function

Function ParseToLy(A As Parse) As String()
ParseToLy = DicToLy(ParseToDic(A))
End Function
Function NewSrcLin(A) As SrcLin
Dim O As New SrcLin
Set NewSrcLin = O.Init(A)
End Function
Function SrcLin_MthDr(A, Lno&, Optional MdNm$) As Variant()
With SrcLin_MthBrk(A)
   SrcLin_MthDr = Array(MdNm, Lno, .Mdy, .Ty, .MthNm)
End With
End Function

Function SrcLin_MthNm$(A)
SrcLin_MthNm = ParseRet(ParseNm(ParseKwMthTy(ParseKwMdy(NewParse(A)))))
End Function

Function SrcLin_IsCd(A) As Boolean
If LinIsEmp(A) Then Exit Function
If SrcLin_IsRmk(A) Then Exit Function
SrcLin_IsCd = True
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

Function SrcLin_TyNm$(A)
SrcLin_TyNm = ParseRet(ParseNm(ParseKwTy(ParseKwMdy(NewParse(A)))))
End Function

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
Function SrcLin_EndLinPfx$(A)
Ass SrcLin_IsMth(A)
SrcLin_EndLinPfx = "End " & LinT1(SrcLin_MthTy(A))
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

Function SrcLin_MthTy$(A)
SrcLin_MthTy = SrcLin_MthBrk(A).Ty
End Function

Property Get SrcLin_EnmNm$(A)
'If SrcLin_IsEmn(A) Then EnmNm = LinNm(NoEnm)
End Property
'
'Function FriendMthLin$()
'If SrcLin_IsMth Then FriendMthLin = "Friend " & NoMdy
'End Function
'
'Function Init(Lin) As SrcLin
'A = Lin
'A_SrcLin_IsMth = HasOneOfPfx(NoMdy, SyOfMthTy)
'Set Init = Me
'End Function
'
'Property Get MthBrk() As MthBrk
'If Not SrcLin_IsMth Then Exit Property
'Dim O As MthBrk
'With O
'    .Mdy = Mdy
'    .MthNm = MthNm
'    .Ty = MthTy
'End With
'MthBrk = O
'End Property
'
'Property Get LinIsEmp() As Boolean
'LinIsEmp = Trim(A) = ""
'End Property
'
'Property Get SrcLin_IsEmn() As Boolean
'SrcLin_IsEmn = HasPfx(NoMdy, C_Enm)
'End Property
'Sub MthBrk__Tst()
'With MthBrk
'    Debug.Print .Mdy
'    Debug.Print .MthNm
'    Debug.Print .Ty
'End With
'End Sub
'Private Function ZZMthLin$()
'ZZMthLin = "Property Get AA()"
'End Function
'Property Get SrcLin_IsMth() As Boolean
'SrcLin_IsMth = A_SrcLin_IsMth
'End Property
'
'Property Get IsPrpLin() As Boolean
'IsPrpLin = HasPfx(NoMdy, C_Prp)
'End Property
'
'Property Get SrcLin_IsRmk() As Boolean
'SrcLin_IsRmk = FstChr(LTrim(A)) = "'"
'End Property
'
'Property Get SrcLin_IsTy() As Boolean
'SrcLin_IsTy = HasPfx(NoMdy, C_Ty)
'End Property
'
'Property Get Mdy$()
'Mdy = StrPfx(A, SyOfMdy)
'End Property
'
'Property Get MthNm$()
'If SrcLin_IsMth Then MthNm = LinNm(NoMthTy)
'End Property
'
'Property Get MthTy$()
'If SrcLin_IsMth Then MthTy = StrPfx(NoMdy, SyOfMthTy)
'End Property
'
'Function NoMdy$()
'NoMdy = LTrim(RmvPfxAy(A, SyOfMdy))
'End Function
'
'Function PrivateMthLin$()
'If SrcLin_IsMth Then PrivateMthLin = "Private " & NoMdy
'End Function
'
'Property Get PrpTy$()
'If IsPrpLin Then PrpTy = LinT1(NoFunTy)
'End Property
'
'Function PrpValDr() As Variant()
'PrpValDr = Array(, , A, EnmNm, LinIsEmp, SrcLin_IsEmn, SrcLin_IsMth, IsPrpLin, SrcLin_IsRmk, SrcLin_IsTy, Mdy, MthNm, MthTy, NoMdy, PrpTy, TyNm)
'End Function
'
'Property Get PrpValFny() As String()
'Static X As Boolean, Y$()
'If Not X Then
'    X = True
'    Y = LvsSy("Md Lno Lin EnmNm LinIsEmp SrcLin_IsEmn SrcLin_IsMth IsPrpLin SrcLin_IsRmk SrcLin_IsTy Mdy MthNm MthTy NoMdy PrpTy TyNm")
'End If
'PrpValFny = Y
'End Property
'
'Function PublicMthLin$()
'If SrcLin_IsMth Then PublicMthLin = NoMdy
'End Function
'
'Property Get TyNm$()
'If SrcLin_IsTy Then TyNm = LinNm(NoTy)
'End Property
'
'Private Property Get NoEnm$()
'If SrcLin_IsEmn Then NoEnm = LTrim(RmvPfx(NoMdy, C_Enm))
'End Property
'
'Private Property Get NoFunTy$()
'If SrcLin_IsMth Then NoFunTy = RmvPfxAy(NoMdy, SyOfFunTy)
'End Property
'
'Private Property Get NoMthTy$()
'If SrcLin_IsMth Then NoMthTy = LTrim(RmvPfxAy(NoMdy, SyOfMthTy))
'End Property
'
'Private Property Get NoTy$()
'If SrcLin_IsTy Then NoTy = LTrim(RmvPfx(NoMdy, C_Ty))
'End Property
'Private Sub AllSrcCode__Tst()
'Dim Dry()
'Dim Dr()
'Dim Drs As Drs
'Dim O$()
'Dim I, Lin
'Dim Md As CodeModule:
'Dim Lno&
'Dim MNm$
'For Each I In PjMbrAy(CurPj)
'    Set Md = I
'    MNm = MdNm(Md)
'    Lno = 0
'    For Each Lin In MdSrc(Md)
'        Lno = Lno + 1
'        A = Lin
'        Dr = PrpValDr
'        Dr(0) = MNm
'        Dr(1) = Lno
'        Push Dry, Dr
'    Next
'Next
'Drs.Dry = Dry
'Drs.Fny = PrpValFny
'DrsWs Drs
'End Sub
'

Private Function ZZSrcLin$()
ZZSrcLin = "Private Sub SrcLin_IsMth()"
End Function
