Attribute VB_Name = "M_BoolOpStr"
Option Explicit
Enum eBoolOp
    eOpEQ = 1
    eOpNE = 2
    eOpAND = 3
    eOpOR = 4
End Enum
Enum eEqNeOp
    eOpEQ = eBoolOp.eOpEQ
    eOpNE = eBoolOp.eOpNE
End Enum

Enum eAndOrOp
    eOpAND = eBoolOp.eOpAND
    eOpOR = eBoolOp.eOpOR
End Enum

Function BoolOpStr_BoolOp(A$) As eBoolOp
Dim O As eBoolOp
Select Case A
Case "AND": O = eBoolOp.eOpAND
Case "OR": O = eBoolOp.eOpOR
Case "EQ": O = eBoolOp.eOpEQ
Case "NE": O = eBoolOp.eOpNE
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
BoolOpStr_IsVdt = IsInUCaseSy(A, SyOf_BoolOp)
End Function
