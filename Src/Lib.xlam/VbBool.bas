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
Type IntOpt
    Int As Integer
    Som As Boolean
End Type
Type BoolOpt
   Bool As Boolean
   Som As Boolean
End Type
Type BoolAyOpt
   Bools As New Bools
   Som As Boolean
End Type

Function BoolAyOpt_And(A As BoolAyOpt) As BoolOpt
If Not A.Som Then Exit Function
BoolAyOpt_And = SomBool(A.Bools.AndVal)
End Function

Function BoolAyOpt_Or(A As BoolAyOpt) As BoolOpt
If Not A.Som Then Exit Function
BoolAyOpt_Or = SomBool(A.Bools.OrVal)
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

Function SomBoolAy(A As Bools) As BoolAyOpt
SomBoolAy.Som = True
Set SomBoolAy.Bools = A
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
    Y = LvsSy("AND OR")
End If
SyOfBoolOp = Y
End Function
