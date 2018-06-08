Attribute VB_Name = "M_Is"
Option Explicit
Function IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Function
Function IsSy(V) As Boolean
IsSy = VarType(V) = vbArray + vbString
End Function
Function IsEmp(V) As Boolean
IsEmp = True
If IsMissing(V) Then Exit Function
If IsNothing(V) Then Exit Function
If IsEmpty(V) Then Exit Function
If IsStr(V) Then
   If V = "" Then Exit Function
End If
If IsArray(V) Then
   If AyIsEmp(V) Then Exit Function
End If
IsEmp = False
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

Function IsBet(V, A, B) As Boolean
If A > V Then Exit Function
If V > B Then Exit Function
IsBet = True
End Function
Property Get IsBool(A) As Boolean
IsBool = VarType(A) = vbBoolean
End Property

Property Get IsDic(A) As Boolean
IsDic = TypeName(A) = "Dictionary"
End Property


Function IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
IsInAp = IsInAy(V, Av)
End Function

Function IsInAy(V, Ay) As Boolean
Dim J%
For J = 0 To UB(Ay)
    If Ay(J) = V Then IsInAy = True: Exit Function
Next
End Function

Function IsInSyIgnCas(V, Sy$()) As Boolean
IsInSyIgnCas = IsInAy(UCase(V), Sy)
End Function

Property Get IsIntAy(V) As Boolean
IsIntAy = VarType(V) = vbArray + vbInteger
End Property

Property Get IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Property

Property Get IsPrim(A) As Boolean
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
End Property

Property Get IsStrAy(A) As Boolean
IsStrAy = VarType(A) = vbArray + vbString
End Property

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





