Attribute VB_Name = "M_BoolAy"
Option Explicit
Public Enum e_BoolAyOp
    e_And = 1
    e_Or = 2
    e_IsAllTrue = 3
    e_IsAllFalse = 4
    e_IsSomTrue = 5
    e_IsSomFalse = 6
End Enum

Property Get BoolAy_AndVal(A() As Boolean) As Boolean
BoolAy_AndVal = BoolAy_IsAllTrue(A)
End Property

Property Get BoolAy_IsAllFalse(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If A(J) Then Exit Property
Next
BoolAy_IsAllFalse = True
End Property

Property Get BoolAy_IsAllTrue(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then Exit Property
Next
BoolAy_IsAllTrue = True
End Property

Property Get BoolAy_IsSomFalse(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If Not A(J) Then BoolAy_IsSomFalse = True: Exit Property
Next
End Property

Property Get BoolAy_IsSomTrue(A() As Boolean) As Boolean
Dim J%
For J = 0 To UB(A)
    If A(J) Then BoolAy_IsSomTrue = True: Exit Property
Next
End Property

Property Get BoolAy_OrVal(A() As Boolean) As Boolean
BoolAy_OrVal = BoolAy_IsSomTrue(A)
End Property

Property Get BoolAy_Val(A() As Boolean, Op As e_BoolAyOp) As Boolean
Dim O As Boolean
Select Case Op
Case e_And, e_IsAllTrue: O = BoolAy_IsAllTrue(A)
Case e_Or, e_IsSomTrue: O = BoolAy_IsSomTrue(A)
Case e_IsAllFalse: O = BoolAy_IsAllFalse(A)
Case e_IsSomFalse: O = BoolAy_IsSomFalse(A)
Case Else: Stop
End Select
BoolAy_Val = O
End Property
