Attribute VB_Name = "VbTpEr"
Option Explicit
Enum e_TpPosTy
    e_PosRCC = 1
    e_PosRR = 2
    e_PosR = 3
End Enum
Type TpPos
    Ty As e_TpPosTy
    R1 As Integer
    R2 As Integer
    C1 As Integer
    C2 As Integer
End Type
Type TpErItm
    Pos As TpPos
    Msg As String
End Type
Type TpEr
    N As Integer
    Ay() As TpErItm
End Type
Type TpErOpt
    Som As Boolean
    Er As TpEr
End Type

Function GpLy(A As Gp) As String()
GpLy = LnxAy_Ly(A.LnxAy)
End Function

Sub GpPush(O() As Gp, M As Gp)
Dim N&
N = GpSz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Function GpSz&(A() As Gp)
On Error Resume Next
GpSz = UBound(A) + 1
End Function

Function GpUB%(A() As Gp)
GpUB = GpSz(A) - 1
End Function

Function NewGp(LnxAy() As Lnx) As Gp
NewGp.LnxAy = LnxAy
End Function

Function NewTpEr(Lx%, Msg$) As TpEr
Dim Ay() As TpErItm
    ReDim Ay(0)
    With Ay(0)
        .Msg = Msg
        With .Pos
            .R1 = Lx
            .Ty = e_PosRCC
        End With
    End With
With NewTpEr
    .N = 1
    .Ay = Ay
End With
End Function

Function NewTpErOfAA() As TpEr
Dim O As TpEr
With O
End With
NewTpErOfAA = O
End Function

Function TpErAp_Add6(A1 As TpEr, A2 As TpEr, A3 As TpEr, A4 As TpEr, A5 As TpEr, A6 As TpEr) As TpEr
Dim Er As TpEr
Er = A1
TpEr_Push Er, A2
TpEr_Push Er, A3
TpEr_Push Er, A4
TpEr_Push Er, A5
TpEr_Push Er, A6
TpErAp_Add6 = Er
End Function

Function TpErItm_FmtStr$(A As TpErItm)
With A
    TpErItm_FmtStr = TpPos_FmtStr(.Pos) & " " & .Msg
End With
End Function

Function TpEr_Add(A As TpEr, B As TpEr) As TpEr
Dim O As TpEr
TpEr_Push O, B
TpEr_Add = O
End Function

Function TpEr_LxAy(A As TpEr) As Integer()
Dim J%, M As TpErItm
Dim Z%(), Ix%
For J = 0 To A.N - 1
    M = A.Ay(J)
    Ix% = M.Pos.R1
    Push Z, Ix
Next
TpEr_LxAy = Z
End Function

Function TpEr_Ly(A As TpEr) As String()
Dim O$(), J%
For J = 0 To A.N - 1
    Push O, TpErItm_FmtStr(A.Ay(J))
Next
TpEr_Ly = O
End Function

Sub TpEr_Push(O As TpEr, A As TpEr)
Dim J%
For J = 0 To A.N - 1
    TpEr_PushItm O, A.Ay(J)
Next
End Sub

Sub TpEr_PushItm(O As TpEr, M As TpErItm)
Dim N%: N = O.N
ReDim Preserve O.Ay(N)
O.Ay(N) = M
O.N = N + 1
End Sub

Sub TpEr_PushLxMsg(O As TpEr, Lx%, Msg$)
TpEr_Push O, NewTpEr(Lx, Msg)
End Sub

Function TpPos_FmtStr$(A As TpPos)
Dim O$
With A
    Select Case .Ty
    Case e_PosRCC
        O = FmtQQ("RCC(? ? ?) ", .R1, .C1, .C2)
    Case e_PosRR
        O = FmtQQ("RR(? ?) ", .R1, .R2)
    Case e_PosR
        O = FmtQQ("R(?)", .R1)
    Case Else
        'Er "TpPos_FmtStr", "Invalid {TpPos}", A.Ty
    End Select
End With
End Function
