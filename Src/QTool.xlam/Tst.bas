Attribute VB_Name = "Tst"
Option Explicit
Public Act, Ept
Function C()
If Not IsEq(Act, Ept) Then Stop
End Function
'
Function IsEq(A, B) As Boolean
Dim T As VbVarType
T = VarType(A)
If T <> VarType(B) Then Exit Function
Select Case True
Case IsArray(A): IsEq = IsEqAy(A, B)
Case IsObject(A): IsEq = ObjPtr(A) = ObjPtr(B)
Case Else: IsEq = A = B
End Select
End Function
Function IsEqAy(A, B) As Boolean
If VarType(A) <> VarType(B) Then Exit Function
Dim U&, J&
U = UB(A)
If UB(B) <> U Then Exit Function
For J = 0 To U
    If Not IsEq(A(J), B(J)) Then
        Debug.Print "IsEqAy: Ele(" & J & ") not equal"
        Debug.Print A(J)
        Debug.Print B(J)
        Exit Function
    End If
Next
IsEqAy = True
End Function
