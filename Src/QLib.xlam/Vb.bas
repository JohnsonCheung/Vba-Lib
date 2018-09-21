Attribute VB_Name = "Vb"
Option Explicit
Type FmTo
    FmIx As Long
    ToIx As Long
End Type
Type RRCC
    R1 As Long
    C1 As Long
    R2 As Long
    C2 As Long
End Type
Type LnoCnt
    Lno As Long
    Cnt As Long
End Type

Property Get ABC(Lin) As ABC
Dim O As New ABC
Set ABC = O.Init(Lin)
End Property

Property Get Bools(A() As Boolean) As Bools
Dim O As New Bools
Set Bools = O.Init(A)
End Property

Property Get C() As Cmd
Static Y As New Cmd
Set C = Y
End Property

Property Get Collx(A As VBA.Collection) As Collx
Dim O As New Collx
Set Collx = O.Init(A)
End Property

Property Get Dix(A As Dictionary) As Dix
Dim O As New Dix
Set Dix = O.Init(A)
End Property

Property Get Emp() As Emp
Static Y As New Emp
Set Emp = Y
End Property

Property Get LABCs() As LABCsBy
Set LABCs = New LABCsBy
End Property

Property Get LABCsRslt(A As LABCs, Optional Er As Er) As LABCsRslt
Dim O As New LABCsRslt
Set LABCsRslt = O.Init(A, Er)
End Property

Property Get Lg() As Logger
Static Y As New Logger
Set Lg = Y
End Property

Property Get Lin(A) As Lin
Dim O As New Lin
Set Lin = O.Init(A)
End Property

Property Get Lines(A) As Lines
Dim O As New Lines
O.Lines = A
Set Lines = O
End Property

Property Get Lnx(Lin$, Lx%) As Lnx1
Dim O As New Lnx1
Set Lnx = O.Init(Lin, Lx)
End Property

Property Get Lnxs(A() As Lnx1) As Lnx1s
Dim O As New Lnx1s
Set Lnxs = O.Init(A)
End Property

Property Get Ly(A$()) As Ly
Dim O As New Ly
Set Ly = O.Init(A)
End Property

Property Get LyRslt(Ly$(), Er As Er) As LyRslt
Dim O As New LyRslt
O.Ly = Ly
Set O.Er = Er
If IsNothing(Er) Then PmEr
Set LyRslt = O
End Property

Property Get Macro(MacroStr$) As Macro
Dim O As New Macro
O.Macro = MacroStr
Set Macro = O
End Property

Property Get Re(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As Re
Dim O As New Re
Set Re = O.Init(Patn, MultiLine, IgnoreCase, IsGlobal)
End Property

Property Get Seed(Seed0) As Seed
Dim O As New Seed
Set Seed = O.Init(Seed0)
End Property

Property Get StrObj(A) As StrObj
Dim O As New StrObj
Set StrObj = O.Init(A)
End Property

Property Get StrRslt(S, Er As Er) As StrRslt
Dim O As New StrRslt
O.Str = S
Set O.Er = Er
If IsNothing(Er) Then PmEr
Set StrRslt = O
End Property

Property Get SyObj(Sy$()) As SyObj
Dim O As New SyObj
Set SyObj = O.Init(Sy)
End Property

Property Get TblNm(A) As TblNm
Dim O As New TblNm
Set TblNm = O.Init(A)
End Property

Property Get TblNms(Ny0) As TblNm
Dim O As New TblNms
Set TblNms = O.Init(Ny0)
End Property

Property Get Tst() As VbTst
Static Y As New VbTst
Set Tst = Y
End Property

Property Get V(A) As V
Dim O As New V
O.Init A
Set V = O
End Property


Sub Asg(V, OV)
If IsObject(V) Then
   Set OV = V
Else
   OV = V
End If
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub

Function AyIsEqSz(A, B) As Boolean
AyIsEqSz = Sz(A) = Sz(B)
End Function

Function CollObjAy(Coll) As Object()
Dim O() As Object
Dim V
For Each V In Coll
   Push O, V
Next
CollObjAy = O
End Function

Function DftTpLy(Tp0) As String()
Select Case True
Case VarIsStr(Tp0): DftTpLy = SplitVBar(Tp0)
Case VarIsSy(Tp0):  DftTpLy = Tp0
Case Else: Stop
End Select
End Function

Sub DtaEr()
MsgBox "DtaEr"
Stop
End Sub

Function ErShow(Er$()) As String()
ErShow = SyShow("Er", Er)
End Function

Function FmToAy_LnoCntAy(A() As FmTo) As LnoCnt()
If FmToAy_VarIsEmp(A) Then Exit Function
Dim U&, J&
    U = FmTo_UB(A)
Dim O() As LnoCnt
   ReDim O(U)
For J = 0 To U
   O(J) = FmTo_LnoCnt(A(J))
Next
FmToAy_LnoCntAy = O
End Function

Function FmToAy_VarIsEmp(A() As FmTo) As Boolean
FmToAy_VarIsEmp = FmTo_Sz(A) = 0
End Function

Function FmTo_HasU(A As FmTo, U&) As Boolean
If U < 0 Then Stop
If IsEmpFmTo(A) Then Exit Function
If A.FmIx > U Then Exit Function
If A.ToIx > U Then Exit Function
FmTo_HasU = True
End Function

Function FmTo_LnoCnt(A As FmTo) As LnoCnt
Dim Lno&, Cnt&
   Cnt = A.ToIx - A.FmIx + 1
   If Cnt < 0 Then Cnt = 0
   Lno = A.FmIx + 1
With FmTo_LnoCnt
   .Cnt = Cnt
   .Lno = Lno
End With
End Function

Function FmTo_N&(A As FmTo)
With A
   FmTo_N = .ToIx - .FmIx + 1
End With
End Function

Sub FmTo_Push(O() As FmTo, M As FmTo)
Dim N&: N = FmTo_Sz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function FmTo_Str$(A As FmTo)
FmTo_Str = FmtQQ("FmTo(? ?)", A.FmIx, A.ToIx)
End Function

Function FmTo_Sz&(A() As FmTo)
On Error Resume Next
FmTo_Sz = UBound(A) + 1
End Function

Function FmTo_UB&(A() As FmTo)
FmTo_UB = FmTo_Sz(A) - 1
End Function

Function IntAyObj(Ay%()) As IntAyObj
Dim O As New IntAyObj
Set IntAyObj = O.Init(Ay)
End Function

Function IsEmpFmTo(A As FmTo) As Boolean
IsEmpFmTo = True
If A.FmIx < 0 Then Exit Function
If A.ToIx < 0 Then Exit Function
If A.FmIx > A.ToIx Then Exit Function
IsEmpFmTo = False
End Function

Function IsEq(Act, Exp) As Boolean
If VarType(Act) <> VarType(Exp) Then Exit Function
If VarIsPrim(Act) Then
    If Act <> Exp Then Exit Function
End If
If IsArray(Act) Then
    If Not AyIsEq(Act, Exp) Then Stop
    Exit Function
End If
End Function

Function IsNonBlankStr(V) As Boolean
If Not IsStr(V) Then Exit Function
IsNonBlankStr = V <> ""
End Function

Function IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
End Function

Function IsNothingOrEmp(V) As Boolean
Select Case TypeName(V)
Case "Nothing", "Empty": IsNothingOrEmp = True
End Select
End Function

Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function

Sub LnoCnt_Dmp(A As LnoCnt)
Debug.Print LnoCnt_Str(A)
End Sub

Sub LnoCnt_Push(O() As LnoCnt, M As LnoCnt)
Dim N&: N = LnoCnt_Sz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function LnoCnt_Str$(A As LnoCnt)
LnoCnt_Str = FmtQQ("Lno(?) Cnt(?)", A.Lno, A.Cnt)
End Function

Function LnoCnt_Sz&(A() As LnoCnt)
On Error Resume Next
LnoCnt_Sz = UBound(A) + 1
End Function

Function LnoCnt_UB&(A() As LnoCnt)
LnoCnt_UB = LnoCnt_Sz(A) - 1
End Function

Function Ly0Ap_Ly(ParamArray Ly0Ap()) As String()
Dim I, Av(): Av = Ly0Ap
If AyIsEmp(Av) Then Exit Function
Dim O$()
For Each I In Av
    PushAy O, DftLy(I)
Next
Ly0Ap_Ly = O
End Function

Function Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) > O Then O = Av(J)
Next
Max = O
End Function

Function Min(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) < O Then O = Av(J)
Next
Min = O
End Function

Sub Never()
Const CSub$ = "Never"
Er CSub, "Should never reach here"
End Sub

Function NewFmTo(FmIx&, ToIx&) As FmTo
NewFmTo.FmIx = FmIx
NewFmTo.ToIx = ToIx
End Function

Function NewLnoCnt(Lno&, Cnt&) As LnoCnt
NewLnoCnt.Lno = Lno
NewLnoCnt.Cnt = Cnt
End Function

Function NewRRCC(R1&, R2&, C1&, C2&) As RRCC
Dim O As RRCC
With O
    .R2 = R2
    .R1 = R1
    .C2 = C2
    .C1 = C1
End With
NewRRCC = O
End Function

Function OkShow(Ok$()) As String()
OkShow = SyShow("Ok", Ok)
End Function

Function Oy(ObjAy) As Oy
Dim O As New Oy
Set Oy = O.Init(ObjAy)
End Function

Sub PmEr()
MsgBox "Parameter Er"
Stop
End Sub

Sub RRCC_Dmp(A As RRCC)
Debug.Print RRCC_Str(A)
End Sub

Function RRCC_IsEmp(A As RRCC) As Boolean
RRCC_IsEmp = True
With A
   If .R1 <= 0 Then Exit Function
   If .R2 <= 0 Then Exit Function
   If .R1 > .R2 Then Exit Function
End With
RRCC_IsEmp = False
End Function

Function RRCC_Str$(A As RRCC)
With A
   RRCC_Str = FmtQQ("(RRCC : ? ? ? ??)", .R1, .R2, .C1, .C2, IIf(RRCC_IsEmp(A), " *Empty", ""))
End With
End Function

Sub Stp()
Stop
End Sub

Function Sy(A$()) As Sy
Dim O As New Sy
Set Sy = O.Init(A)
End Function

Function SyShow(XX$, Sy$()) As String()
Dim O$()
Select Case Sz(Sy)
Case 0
    Push O, XX & "()"
Case 1
    Push O, XX & "(" & Sy(0) & ")"
Case Else
    Push O, XX & "("
    PushAy O, Sy
    Push O, XX & ")"
End Select
SyShow = O
End Function

Function Tag$(TagNm$, S)
If HasPfx(S, TagNm & "(") Then
    If HasSfx(S, ")") Then
        Tag = S
        Exit Function
    End If
End If
If Has(S, vbCrLf) Then
    Tag = FmtQQ("?(|?|?)", TagNm, S, TagNm)
Else
    Tag = FmtQQ("?(?)", TagNm, S)
End If
End Function

Function Tag_NyStr_ObjAp$(TagNm$, NyStr$, ParamArray ObjAp())
Dim Av(): Av = ObjAp
Tag_NyStr_ObjAp = Tag_Ny_ObjAv(TagNm, LvsSy(NyStr), Av)
End Function

Function VarIsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
VarIsBet = True
End Function

Function VarIsBool(V) As Boolean
VarIsBool = VarType(V) = vbBoolean
End Function

Function VarIsDic(A) As Boolean
VarIsDic = TypeName(A) = "Dictionary"
End Function

Function VarIsEmp(V) As Boolean
VarIsEmp = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsEmpty(V) Then Exit Function
If VarIsStr(V) Then
   If V = "" Then Exit Function
End If
If IsArray(V) Then
   If AyIsEmp(V) Then Exit Function
End If
VarIsEmp = False
End Function

Function VarIsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
VarIsInAp = AyHas(Av, V)
End Function

Function VarIsInAy(V, Ay) As Boolean
VarIsInAy = AyHas(Ay, V)
End Function

Function VarIsInUcaseSy(V, Sy$()) As Boolean
VarIsInUcaseSy = VarIsInAy(UCase(V), Sy)
End Function

Function VarIsIntAy(V) As Boolean
VarIsIntAy = VarType(V) = vbArray + vbInteger
End Function

Function VarIsLngAy(V) As Boolean
VarIsLngAy = VarType(V) = vbArray + vbLong
End Function

Function VarIsPrim(V) As Boolean
Select Case VarType(V)
Case _
   VbVarType.vbBoolean, _
   VbVarType.vbByte, _
   VbVarType.vbCurrency, _
   VbVarType.vbDate, _
   VbVarType.vbDecimal, _
   VbVarType.vbDouble, _
   VbVarType.vbInteger, _
   VbVarType.vbLong, _
   VbVarType.vbSingle, _
   VbVarType.vbString
   VarIsPrim = True
End Select
End Function

Function VarIsStr(V) As Boolean
VarIsStr = VarType(V) = vbString
End Function

Function VarIsStrAy(V) As Boolean
VarIsStrAy = VarType(V) = vbArray + vbString
End Function

Function VarIsSy(V) As Boolean
VarIsSy = VarIsStrAy(V)
End Function

Private Sub VarIsStrAy__Tst()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Ass VarIsStrAy(A) = True
Ass VarIsStrAy(B) = True
Ass VarIsStrAy(C) = False
Ass VarIsStrAy(D) = False
End Sub

Private Function Tag_Ny_ObjAv$(TagNm$, Ny$(), ObjAv())
Ass AyIsSamSz(Ny, ObjAv)
Dim S$
    Dim O$()
    Dim A$, N%
    Dim J%
    For J = 0 To UB(Ny)
        Select Case True
        Case IsNothing(ObjAv(J)): A = "Nothing"
        Case IsEmpty(ObjAv(J)):   A = "Empty"
        Case Else:                A = CallByName(ObjAv(J), "ToStr", VbGet)
        End Select
        Push O, Tag(Ny(J), A)
    Next
    S = JnCrLf(O)
Tag_Ny_ObjAv = Tag(TagNm, S)
End Function
