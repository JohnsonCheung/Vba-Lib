Attribute VB_Name = "M_Is"
Option Explicit

Function IsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
IsBet = True
End Function

Function IsBool(A) As Boolean
IsBool = VarType(A) = vbBoolean
End Function

Function IsDic(A) As Boolean
IsDic = TypeName(A) = "Dictionary"
End Function

Function IsDigit(A) As Boolean
IsDigit = "0" <= A And A <= "9"
End Function

Function IsEmp(V) As Boolean
IsEmp = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsEmpty(V) Then Exit Function
If IsStr(V) Then
   If V = "" Then Exit Function
End If
End Function

Function IsEq(Act, Exp) As Boolean
'If VarType(Act) <> VarType(Exp) Then Exit Function
'If IsPrim(Act) Then
'    If Act <> Exp Then Exit Function
'End If
'If IsArray(Act) Then
'    If Not AyIsEq(Act, Exp) Then Stop
'    Exit Function
'End If
End Function

Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
IsInAp = AyHas(Av, V)
End Function

Function IsInUCaseSy(V, Sy$()) As Boolean
IsInUCaseSy = AyHas(Sy, UCase(V))
End Function

Function IsIntAy(V) As Boolean
IsIntAy = VarType(V) = vbArray + vbInteger
End Function

Function IsLetter(A) As Boolean
Dim C1$: C1 = UCase(A)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Function

Function IsLng(A) As Boolean
IsLng = VarType(A) = vbLong
End Function

Function IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Function

Function IsNeedQuote(A) As Boolean
IsNeedQuote = True
If HasSubStr(A, " ") Then Exit Function
If HasSubStr(A, "#") Then Exit Function
If HasSubStr(A, ".") Then Exit Function
IsNeedQuote = False
End Function

Function IsNm(A) As Boolean
If Not IsLetter(FstChr(A)) Then Exit Function
Dim L%: L = Len(A)
If L > 64 Then Exit Function
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(A, J, 1)) Then Exit Function
Next
IsNm = True
End Function

Function IsNmChr(A) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Function
If A = "_" Then Exit Function
If IsDigit(A) Then Exit Function
IsNmChr = False
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

Function IsPrim(A) As Boolean
Select Case VarType(A)
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
   IsPrim = True
End Select
End Function

Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function

Function IsStrAy(A) As Boolean
IsStrAy = VarType(A) = vbArray + vbString
End Function

Function IsSy(V) As Boolean
IsSy = VarType(V) = vbArray + vbString
End Function

Function IsVdtVbl(A) As Boolean
If Not IsStr(A) Then Exit Function
If HasSubStr(A, vbCr) Then Exit Function
If HasSubStr(A, vbLf) Then Exit Function
IsVdtVbl = True
End Function

Function IsWhite(A) As Boolean
Dim B$: B = Left(A, 1)
IsWhite = True
If B = " " Then Exit Function
If B = vbCr Then Exit Function
If B = vbLf Then Exit Function
If B = vbTab Then Exit Function
IsWhite = False
End Function

Sub ZZ__Tst()
ZZ_IsStrAy
End Sub

Private Sub ZZ_IsStrAy()
Dim A$()
Dim B: B = A
Dim C()
Dim D
Ass IsStrAy(A) = True
Ass IsStrAy(B) = True
Ass IsStrAy(C) = False
Ass IsStrAy(D) = False
End Sub
