Attribute VB_Name = "M_BoolAy"
Option Explicit
Public Enum eBoolAyOp
    eAnd = 1
    eOr = 2
    eIsAllTrue = 3
    eIsAllFalse = 4
    eIsSomTrue = 5
    eIsSomFalse = 6
End Enum

Function BoolAy_AndVal(A() As Boolean) As Boolean
BoolAy_AndVal = BoolAy_IsAllTrue(A)
End Function

Function BoolAy_IsAllFalse(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If A(J) Then Exit Function
Next
BoolAy_IsAllFalse = True
End Function

Function BoolAy_IsAllTrue(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then Exit Function
Next
BoolAy_IsAllTrue = True
End Function

Function BoolAy_IsSomFalse(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then BoolAy_IsSomFalse = True: Exit Function
Next
End Function

Function BoolAy_IsSomTrue(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If A(J) Then BoolAy_IsSomTrue = True: Exit Function
Next
End Function

Function BoolAy_OrVal(A() As Boolean) As Boolean
BoolAy_OrVal = BoolAy_IsSomTrue(A)
End Function

Function BoolAy_Val(A() As Boolean, Op As eBoolAyOp) As Boolean
Dim O As Boolean
Select Case Op
Case eAnd, eIsAllTrue: O = BoolAy_IsAllTrue(A)
Case eOr, eIsSomTrue: O = BoolAy_IsSomTrue(A)
Case eIsAllFalse: O = BoolAy_IsAllFalse(A)
Case eIsSomFalse: O = BoolAy_IsSomFalse(A)
Case Else: Stop
End Select
End Function
