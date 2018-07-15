Attribute VB_Name = "M_Parse"
Option Explicit
Type Parse
    Lin As String
    IsOk As Boolean
    Ok As String
End Type

Property Get ParseOneTerm(A, TermAy$())  'As Parse
If Not A.IsOk Then ParseOneTerm = A: Exit Property
Dim F$: F = StrPfx(A.Lin, TermAy)
If F = "" Then
   Dim Msg$: Msg = FmtQQ("These Terms[?] not found", JnVBar(TermAy))
   ParseOneTerm = NewErParse(Msg, A.Lin)
Else
   ParseOneTerm = NewOkParse(F, LTrim(RmvPfx(A.Lin, F)))
End If
End Property

Property Get Parse_seKwBktPair(A As Parse) As Parse
ParseKwBktPair = ParseStr(A, "()")
End Property

Property Get Parse_seKwEnm(A As Parse) As Parse
ParseKwEnm = ParseTerm(A, "Enum")
End Property

Property Get Parse_seKwMdy(A As Parse) As Parse
ParseKwMdy = ParseOptOneTerm(A, SyOfMdy)
End Property

Property Get Parse_seKwMthTy(A As Parse) As Parse
ParseKwMthTy = ParseOneTerm(A, SyOfMthTy)
End Property

Property Get Parse_seKwOptBktPair(A As Parse) As Parse
ParseKwOptBktPair = ParseOpt(ParseKwBktPair(A))
End Property

Property Get Parse_seKwOptional(A As Parse) As Parse
ParseKwOptional = ParseTerm(A, "Optional")
End Property

Property Get Parse_seKwPrmAy(A As Parse) As Parse
ParseKwPrmAy = ParseTerm(A, "ParamArray")
End Property

Property Get Parse_seKwTy(A As Parse) As Parse
ParseKwTy = ParseTerm(A, "Type")
End Property

Property Get Parse_seKwTyChr(A As Parse) As Parse
ParseKwTyChr = ParseOptOneChr(A, TyChrLis)
End Property

Property Get Parse_seNm(A As Parse) As Parse
If Not A.IsOk Then ParseNm = A: Exit Property
Dim B$
Stop
'   B = Lin(A.Lin).Nm

Dim L&: L = Len(B)
If L = 0 Then
   ParseNm = NewErParse("No name", A.Lin)
Else
   ParseNm = NewOkParse(B, Mid(A.Lin, L + 1))
End If
End Property

Property Get Parse_seOneChr(A As Parse, ChrLis$) As Parse
If Not A.IsOk Then ParseOneChr = A: Exit Property
Dim C$: C = FstChr(A.Lin)
If HasSubStr(ChrLis, C) Then
   ParseOneChr = NewOkParse(C, RmvFstChr(A.Lin))
Else
   ParseOneChr = NewErParse(FmtQQ("One of ChrLis[?] not found", ChrLis), A.Lin)
End If
End Property

Property Get Parse_seOneTerm(A As Parse, TermAy$()) As Parse
Dim O As Parse
Dim J%
For J = 0 To UB(TermAy)
   O = ParseTerm(A, TermAy(J))
   If O.IsOk Then ParseOneTerm = O: Exit Property
Next
Dim Msg$
   Msg = FmtQQ("These Term[?] not found", JnSpc(TermAy))
ParseOneTerm = NewErParse(Msg, A.Lin)
End Property

Property Get Parse_seOpt(A As Parse) As Parse
If A.IsOk Then ParseOpt = A: Exit Property
A.IsOk = True
A.Er_or_Ok = ""
ParseOpt = A
End Property

Property Get Parse_seOptOneChr(A As Parse, ChrLis$) As Parse
ParseOptOneChr = ParseOpt(ParseOneChr(A, ChrLis))
End Property

Property Get Parse_seOptOneTerm(A As Parse, TermAy$()) As Parse
ParseOptOneTerm = ParseOpt(ParseOneTerm(A, TermAy))
End Property

Property Get Parse_seRet$(A As Parse)
If A.IsOk Then ParseRet = A.Er_or_Ok
End Property

Property Get Parse_seRmvSpc(A As Parse) As Parse
If Not A.IsOk Then ParseRmvSpc = A: Exit Property
A.Lin = LTrim(A.Lin)
ParseRmvSpc = A
End Property

Property Get Parse_seStr(A As Parse, Str$) As Parse
If Not A.IsOk Then ParseStr = A: Exit Property
If Not HasPfx(A.Lin, Str) Then ParseStr = NewErParse(FmtQQ("[?] not found", Str), A.Lin): Exit Property
ParseStr = NewOkParse(Str, RmvPfx(A.Lin, Str))
End Property

Property Get Parse_seTerm(A As Parse, Term$) As Parse
ParseTerm = ParseRmvSpc(ParseStr(A, Term))
End Property

Property Get Parse_seToDic(A As Parse) As Dictionary
Dim O As New Dictionary
With O
   .Add "NewParse", A.Lin
   .Add "Is", IIf(A.IsOk, "Ok", "Er")
   .Add IIf(A.IsOk, "Rslt", "Er"), A.Er_or_Ok
End With
Set ParseToDic = O
End Property

Property Get Parse_seToLy(A As Parse) As String()
ParseToLy = Dix(ParseToDic(A)).Ly
End Property

Sub ParseBrw(A As Parse)
aybrw ParseToLy(A)
End Sub

Sub ParseDmp(A As Parse)
AyDmp ParseToLy(A)
End Sub
