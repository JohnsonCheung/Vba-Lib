Attribute VB_Name = "M_Bool"
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

Property Get BoolOpStr_BoolOp(A$) As e_BoolOp
Dim O As e_BoolOp
Select Case A
Case "AND": O = e_BoolOp.e_OpAND
Case "OR": O = e_BoolOp.e_OpOR
Case "EQ": O = e_BoolOp.e_OpEQ
Case "NE": O = e_BoolOp.e_OpNE
End Select
BoolOpStr_BoolOp = O
End Property

Property Get BoolOpStr_IsAndOr(A$) As Boolean
Select Case UCase(A)
Case "AND", "OR": BoolOpStr_IsAndOr = True
End Select
End Property

Property Get BoolOpStr_IsEqNe(A$) As Boolean
Select Case UCase(A)
Case "EQ", "NE": BoolOpStr_IsEqNe = True
End Select
End Property

Property Get BoolOpStr_IsVdt(A$) As Boolean
BoolOpStr_IsVdt = IsInUCaseSy(A, SyOfBoolOp)
End Property

Property Get SyOfBoolOp() As String()
Static Y$(), X As Boolean
If Not X Then
    X = True
    Y = SslSy("AND OR")
End If
SyOfBoolOp = Y
End Property
