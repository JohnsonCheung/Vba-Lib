Attribute VB_Name = "Vb"
Option Explicit
Public Const M_Val_IsNonNum$ = "Lx(?) has Val(?) should be a number"
Public Const M_Val_ShouldBet$ = "Lx(?) has Val(?) should be between [?] and [?]"
Public Const M_Fld_IsInValid$ = "Lx(?) Fld(?) is invalid.  Not found in Fny"
Public Const M_Fld_IsDup$ = "Lx(?) Fld(?) is found dup in Lx(?)."
Type FmtO
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
Property Get Lines(A) As Lines
Dim O As New Lines
O.Lines = A
Set Lines = O
End Property
Property Get Lin(A) As Lin
Dim O As New Lin
O.Lin = A
Set Lin = O
End Property
Property Get ABC(Lin) As ABC
Dim O As New ABC
Set ABC = O.Init(Lin)
End Property

Function Oy(ObjAy) As Oy
Dim O As New Oy
Set Oy = O.Init(ObjAy)
End Function

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

Function CmpTy_Str$(A As vbext_ComponentType)
Dim O$
Select Case A
Case vbext_ComponentType.vbext_ct_ActiveXDesigner: O = "ActiveXDesigner"
Case vbext_ComponentType.vbext_ct_ClassModule: O = "Class"
Case vbext_ComponentType.vbext_ct_Document: O = "Doc"
Case vbext_ComponentType.vbext_ct_MSForm: O = "MsForm"
Case vbext_ComponentType.vbext_ct_StdModule: O = "Md"
Case Else: O = "Unknown(" & A & ")"
End Select
CmpTy_Str = O
End Function

Function CollObjAy(Coll) As Object()
Dim O() As Object
Dim V
For Each V In Coll
   Push O, V
Next
CollObjAy = O
End Function

Function FmToAy_LnoCntAy(A() As FmtO) As LnoCnt()
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

Function FmToAy_VarIsEmp(A() As FmtO) As Boolean
FmToAy_VarIsEmp = FmTo_Sz(A) = 0
End Function

Function FmTo_HasU(A As FmtO, U&) As Boolean
If U < 0 Then Stop
If IsEmpFmTo(A) Then Exit Function
If A.FmIx > U Then Exit Function
If A.ToIx > U Then Exit Function
FmTo_HasU = True
End Function

Function FmTo_LnoCnt(A As FmtO) As LnoCnt
Dim Lno&, Cnt&
   Cnt = A.ToIx - A.FmIx + 1
   If Cnt < 0 Then Cnt = 0
   Lno = A.FmIx + 1
With FmTo_LnoCnt
   .Cnt = Cnt
   .Lno = Lno
End With
End Function

Function FmTo_N&(A As FmtO)
With A
   FmTo_N = .ToIx - .FmIx + 1
End With
End Function

Sub FmTo_Push(O() As FmtO, M As FmtO)
Dim N&: N = FmTo_Sz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function FmTo_Str$(A As FmtO)
FmTo_Str = FmtQQ("FmTo(? ?)", A.FmIx, A.ToIx)
End Function

Function FmTo_Sz&(A() As FmtO)
On Error Resume Next
FmTo_Sz = UBound(A) + 1
End Function

Function FmTo_UB&(A() As FmtO)
FmTo_UB = FmTo_Sz(A) - 1
End Function

Function IsEmpFmTo(A As FmtO) As Boolean
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

Function IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
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

Function NewFmTo(FmIx&, ToIx&) As FmtO
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
Function AyIsEqSz(A, B) As Boolean
AyIsEqSz = Sz(A) = Sz(B)
End Function
Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function
Property Get Cmd() As Cmd
Static Y As New Cmd
Set Cmd = Y
End Property
Property Get V(A) As V
Dim O As New V
O.Init A
Set V = O
End Property
Property Get LABCsRslt(A As LABCs, Optional Er As Er) As LABCsRslt
Dim O As New LABCsRslt
Set LABCsRslt = O.Init(A, Er)
End Property
Function StrAp_Lines(ParamArray StrAp())
Dim I, Av(): Av = StrAp
If AyIsEmp(Av) Then Exit Function
Dim O$()
For Each I In Av
    If I <> "" Then Push O, I
Next
StrAp_Lines = JnCrLf(O)
End Function

Property Get Coll(A As VBA.Collection) As Coll
Dim O As New Coll
Set Coll = O.Init(A)
End Property
Property Get LABCs() As LABCsBy
Set LABCs = New LABCsBy
End Property
Function SrcLin(A) As SrcLin
Dim O As New SrcLin
O.Init A
Set SrcLin = O
End Function
Sub DtaEr()
MsgBox "DtaEr"
Stop
End Sub
Function Sy(A$()) As Sy
Dim O As New Sy
Set Sy = O.Init(A)
End Function
Function ErShow(Er$()) As String()
ErShow = SyShow("Er", Er)
End Function
Function OkShow(Ok$()) As String()
OkShow = SyShow("Ok", Ok)
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
Sub PrmEr()
MsgBox "Prm Er"
Stop
End Sub

