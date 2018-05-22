Attribute VB_Name = "VbBool"
Option Explicit
Enum e_BoolOp
    e_OpEQ = 1
    e_OpNE = 2
    e_OpAND = 3
    e_OpOR = 4
End Enum
Enum e_EqNeOp
    e_OpEQ = e_BoolOp.e_OpEQ
    e_OpNE = e_BoolOp.e_OpNE
End Enum

Enum e_AndOrOp
    e_OpAND = e_BoolOp.e_OpAND
    e_OpOR = e_BoolOp.e_OpOR
End Enum
Enum e_BoolAyOp
    e_And = 1
    e_Or = 2
End Enum
Type IntOpt
    Int As Integer
    Som As Boolean
End Type
Type BoolOpt
   Bool As Boolean
   Som As Boolean
End Type
Type BoolAyOpt
   BoolAy() As Boolean
   Som As Boolean
End Type

Function BoolAyOpt_And(A As BoolAyOpt) As BoolOpt
If Not A.Som Then Exit Function
BoolAyOpt_And = SomBool(BoolAy_And(A.BoolAy))
End Function

Function BoolAyOpt_Or(A As BoolAyOpt) As BoolOpt
If Not A.Som Then Exit Function
BoolAyOpt_Or = SomBool(False)
End Function

Function BoolAy_And(A() As Boolean) As Boolean
Dim I
For Each I In A
   If Not I Then Exit Function
Next
BoolAy_And = True
End Function

Function BoolAy_Or(A() As Boolean) As Boolean
Dim I
If AyIsEmp(A) Then Exit Function
For Each I In A
   If I Then BoolAy_Or = True: Exit Function
Next
End Function

Function BoolAy_Val(A() As Boolean, Op As e_BoolAyOp) As Boolean
Select Case Op
Case e_BoolAyOp.e_And: BoolAy_Val = BoolAy_And(A)
Case e_BoolAyOp.e_Or: BoolAy_Val = BoolAy_Or(A)
Case Else: Stop
End Select
End Function

Function BoolOpStr_BoolOp(A$) As e_BoolOp
Dim O As e_BoolOp
Select Case A
Case "AND": O = e_BoolOp.e_OpAND
Case "OR": O = e_BoolOp.e_OpOR
Case "EQ": O = e_BoolOp.e_OpEQ
Case "NE": O = e_BoolOp.e_OpNE
End Select
BoolOpStr_BoolOp = O
End Function

Function BoolOpStr_IsAndOr(A$) As Boolean
Select Case UCase(A)
Case "AND", "OR": BoolOpStr_IsAndOr = True
End Select
End Function

Function BoolOpStr_IsEqNe(A$) As Boolean
Select Case UCase(A)
Case "EQ", "NE": BoolOpStr_IsEqNe = True
End Select
End Function

Function BoolOpStr_IsVdt(A$) As Boolean
BoolOpStr_IsVdt = VarIsInUcaseSy(A, SyOfBoolOp)
End Function

Function SomBool(Bool) As BoolOpt
SomBool.Som = True
SomBool.Bool = Bool
End Function

Function SomBoolAy(A() As Boolean) As BoolAyOpt
SomBoolAy.Som = True
SomBoolAy.BoolAy = A
End Function

Function SomInt(I%) As IntOpt
With SomInt
    .Int = I
    .Som = True
End With
End Function

Function SyOfBoolOp() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = LvsSy("AND OR EQ NE")
End If
SyOfBoolOp = Y
End Function
