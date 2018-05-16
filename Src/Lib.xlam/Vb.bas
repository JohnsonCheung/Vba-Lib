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

Function IsEq(Act, Exp) As Boolean
If VarType(Act) <> VarType(Exp) Then Exit Function
If ValIsPrim(Act) Then
    If Act <> Exp Then Exit Function
End If
If IsArray(Act) Then
    If Not AyIsEq(Act, Exp) Then Stop
    Exit Function
End If
End Function


Function ValIsPrim(V) As Boolean
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
   ValIsPrim = True
End Select
End Function

Function ValIsStr(V) As Boolean
ValIsStr = VarType(V) = vbString
End Function

Function ValIsStrAy(V) As Boolean
ValIsStrAy = VarType(V) = vbArray + vbString
End Function

Function ValIsSy(V) As Boolean
ValIsSy = ValIsStrAy(V)
End Function
Function ValIsIntAy(V) As Boolean
ValIsIntAy = VarType(V) = vbArray + vbInteger
End Function

Function ValIsLngAy(V) As Boolean
ValIsLngAy = VarType(V) = vbArray + vbLong
End Function

Function IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
End Function

Function ValIsBool(V) As Boolean
ValIsBool = VarType(V) = vbBoolean
End Function

Function RRCC_IsEmp(A As RRCC) As Boolean
RRCC_IsEmp = True
With A
   If .R1 <= 0 Then Exit Function
   If .R2 <= 0 Then Exit Function
   If .R1 > .R2 Then Exit Function
End With
RRCC_IsEmp = False
End Function


Function ValIsEmp(V) As Boolean
ValIsEmp = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsEmpty(V) Then Exit Function
If ValIsStr(V) Then
   If V = "" Then Exit Function
End If
If IsArray(V) Then
   If AyIsEmp(V) Then Exit Function
End If
ValIsEmp = False
End Function

Function ValIsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
ValIsInAp = AyHas(Av, V)
End Function
Function ValIsInAy(V, Ay) As Boolean
ValIsInAy = AyHas(Ay, V)
End Function
Function ValIsInUcaseSy(V, Sy$()) As Boolean
ValIsInUcaseSy = ValIsInAy(UCase(V), Sy)
End Function

Private Sub ValIsStrAy__Tst()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Ass ValIsStrAy(A) = True
Ass ValIsStrAy(B) = True
Ass ValIsStrAy(C) = False
Ass ValIsStrAy(D) = False
End Sub



Function NewLnoCnt(Lno&, Cnt&) As LnoCnt
NewLnoCnt.Lno = Lno
NewLnoCnt.Cnt = Cnt
End Function

Sub LnoCnt_Dmp(A As LnoCnt)
Debug.Print LnoCnt_Str(A)
End Sub

Sub LnoCnt_Push(O() As LnoCnt, M As LnoCnt)
Dim N&: N = LnoCnt_Sz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function LnoCnt_Sz&(A() As LnoCnt)
On Error Resume Next
LnoCnt_Sz = UBound(A) + 1
End Function

Function LnoCnt_Str$(A As LnoCnt)
LnoCnt_Str = FmtQQ("Lno(?) Cnt(?)", A.Lno, A.Cnt)
End Function

Function LnoCnt_UB&(A() As LnoCnt)
LnoCnt_UB = LnoCnt_Sz(A) - 1
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

Function RRCC_Str$(A As RRCC)
With A
   RRCC_Str = FmtQQ("(RRCC : ? ? ? ??)", .R1, .R2, .C1, .C2, IIf(RRCC_IsEmp(A), " *Empty", ""))
End With
End Function

Function FmToAy_LnoCntAy(A() As FmTo) As LnoCnt()
If FmToAy_ValIsEmp(A) Then Exit Function
Dim U&, J&
    U = FmTo_UB(A)
Dim O() As LnoCnt
   ReDim O(U)
For J = 0 To U
   O(J) = FmTo_LnoCnt(A(J))
Next
FmToAy_LnoCntAy = O
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

Function FmTo_Sz&(A() As FmTo)
On Error Resume Next
FmTo_Sz = UBound(A) + 1
End Function

Function FmTo_Str$(A As FmTo)
FmTo_Str = FmtQQ("FmTo(? ?)", A.FmIx, A.ToIx)
End Function

Function FmTo_UB&(A() As FmTo)
FmTo_UB = FmTo_Sz(A) - 1
End Function

Function IsEmpFmTo(A As FmTo) As Boolean
IsEmpFmTo = True
If A.FmIx < 0 Then Exit Function
If A.ToIx < 0 Then Exit Function
If A.FmIx > A.ToIx Then Exit Function
IsEmpFmTo = False
End Function

Function FmToAy_ValIsEmp(A() As FmTo) As Boolean
FmToAy_ValIsEmp = FmTo_Sz(A) = 0
End Function

Function FmTo_HasU(A As FmTo, U&) As Boolean
If U < 0 Then Stop
If IsEmpFmTo(A) Then Exit Function
If A.FmIx > U Then Exit Function
If A.ToIx > U Then Exit Function
FmTo_HasU = True
End Function

Function NewFmTo(FmIx&, ToIx&) As FmTo
NewFmTo.FmIx = FmIx
NewFmTo.ToIx = ToIx
End Function

Function ValIsDic(A) As Boolean
ValIsDic = TypeName(A) = "Dictionary"
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
Function ValIsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
ValIsBet = True
End Function

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

