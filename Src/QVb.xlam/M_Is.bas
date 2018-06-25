Attribute VB_Name = "M_Is"
Option Explicit

Property Get IsBet(V, A, B) As Boolean
If A > V Then Exit Property
If V > B Then Exit Property
IsBet = True
End Property

Property Get IsBool(A) As Boolean
IsBool = VarType(A) = vbBoolean
End Property

Property Get IsDic(A) As Boolean
IsDic = TypeName(A) = "Dictionary"
End Property

Property Get IsDigit(A) As Boolean
IsDigit = "0" <= A And A <= "9"
End Property

Property Get IsEmp(V) As Boolean
IsEmp = True
If IsMissing(V) Then Exit Property
If IsNothing(V) Then Exit Property
If IsEmpty(V) Then Exit Property
If IsStr(V) Then
   If V = "" Then Exit Property
End If
If IsArray(V) Then
   If AyIsEmp(V) Then Exit Property
End If
IsEmp = False
End Property

Property Get IsInAp(V, ParamArray Ap()) As Boolean
Dim Av(): Av = Ap
IsInAp = IsInAy(V, Av)
End Property

Property Get IsInAy(V, Ay) As Boolean
Dim J%
For J = 0 To UB(Ay)
    If Ay(J) = V Then IsInAy = True: Exit Property
Next
End Property

Property Get IsInUCaseSy(V, Sy$()) As Boolean
IsInUCaseSy = IsInAy(UCase(V), Sy)
End Property

Property Get IsIntAy(V) As Boolean
IsIntAy = VarType(V) = vbArray + vbInteger
End Property

Property Get IsLetter(A) As Boolean
Dim C1$: C1 = UCase(A)
IsLetter = ("A" <= C1 And C1 <= "Z")
End Property

Property Get IsLng(A) As Boolean
IsLng = VarType(A) = vbLong
End Property

Property Get IsLngAy(V) As Boolean
IsLngAy = VarType(V) = vbArray + vbLong
End Property

Property Get IsNeedQuote(A) As Boolean
IsNeedQuote = True
If HasSubStr(A, " ") Then Exit Property
If HasSubStr(A, "#") Then Exit Property
If HasSubStr(A, ".") Then Exit Property
IsNeedQuote = False
End Property

Property Get IsNm(A) As Boolean
If Not IsLetter(FstChr(A)) Then Exit Property
Dim L%: L = Len(A)
If L > 64 Then Exit Property
Dim J%
For J = 2 To L
   If Not IsNmChr(Mid(A, J, 1)) Then Exit Property
Next
IsNm = True
End Property

Property Get IsNmChr(A) As Boolean
IsNmChr = True
If IsLetter(A) Then Exit Property
If A = "_" Then Exit Property
If IsDigit(A) Then Exit Property
IsNmChr = False
End Property

Property Get IsNonBlankStr(V) As Boolean
If Not IsStr(V) Then Exit Property
IsNonBlankStr = V <> ""
End Property

Property Get IsNothing(V) As Boolean
IsNothing = TypeName(V) = "Nothing"
End Property

Property Get IsNothingOrEmp(V) As Boolean
Select Case TypeName(V)
Case "Nothing", "Empty": IsNothingOrEmp = True
End Select
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

Property Get IsStr(V) As Boolean
IsStr = VarType(V) = vbString
End Property

Property Get IsStrAy(A) As Boolean
IsStrAy = VarType(A) = vbArray + vbString
End Property

Property Get IsSy(V) As Boolean
IsSy = VarType(V) = vbArray + vbString
End Property

Property Get IsVdtVbl(A) As Boolean
If Not IsStr(A) Then Exit Property
If HasSubStr(A, vbCr) Then Exit Property
If HasSubStr(A, vbLf) Then Exit Property
IsVdtVbl = True
End Property

Property Get IsWhite(A) As Boolean
Dim B$: B = Left(A, 1)
IsWhite = True
If B = " " Then Exit Property
If B = vbCr Then Exit Property
If B = vbLf Then Exit Property
If B = vbTab Then Exit Property
IsWhite = False
End Property

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
