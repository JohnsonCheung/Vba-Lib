Attribute VB_Name = "SalRpt__SqlOf_TTx__Tst"
Option Explicit
Private Type ZZ6_TstDta
   P          As SrPm
   ShouldThow As Boolean
   Exp        As String
End Type

Private Function Act$(A As ZZ6_TstDta)
Act = Srp_TTx(A.P)
End Function

Private Function ActOpt(A As ZZ6_TstDta) As SomStr
On Error GoTo X
With A
   ActOpt = SomStr(Act(A))
End With
Exit Function
X:
End Function

Private Sub ZZ6_Push(O() As ZZ6_TstDta, I As ZZ6_TstDta)
Dim N&: N = ZZ6_Sz(O)
ReDim Preserve O(N)
O(N) = I
End Sub

Private Function ZZ6_Sz%(A() As ZZ6_TstDta)
On Error Resume Next
ZZ6_Sz = UBound(A) + 1
End Function

Private Function ZZ6_TstDta0() As ZZ6_TstDta
With ZZ6_TstDta0
   With .P
   End With
   .ShouldThow = False
   .Exp = ""
End With
End Function

Private Function ZZ6_TstDta1() As ZZ6_TstDta
With ZZ6_TstDta1
   With .P
   End With
   .ShouldThow = False
   .Exp = ""
End With
End Function

Private Function ZZ6_TstDta2() As ZZ6_TstDta
With ZZ6_TstDta2
   With .P
   End With
   .ShouldThow = False
   .Exp = ""
End With
End Function

Private Function ZZ6_TstDta3() As ZZ6_TstDta
With ZZ6_TstDta3
   With .P
   End With
   .ShouldThow = False
   .Exp = ""
End With
End Function

Private Function ZZ6_TstDtaAy() As ZZ6_TstDta()
Dim O() As ZZ6_TstDta
ZZ6_Push O, ZZ6_TstDta0
ZZ6_Push O, ZZ6_TstDta1
ZZ6_Push O, ZZ6_TstDta2
ZZ6_Push O, ZZ6_TstDta3
ZZ6_TstDtaAy = O
End Function

Private Sub ZZ6_Tstr(A As ZZ6_TstDta)
Dim M As SomStr
   M = ActOpt(A)
With A
   If .ShouldThow Then
       If M.Som Then Stop
   Else
       If Not M.Som Then Stop
       Ass IsEq(M.Str, .Exp)
   End If
End With
End Sub

Private Function ZZ6_UB%(A() As ZZ6_TstDta)
ZZ6_UB = ZZ6_Sz(A) - 1
End Function

Sub Srp_TTx__Tst()
Dim J%
Dim Ay() As ZZ6_TstDta: Ay = ZZ6_TstDtaAy
For J = 0 To ZZ6_UB(Ay)
   If J = J Then
       ZZ6_Tstr Ay(J)
   End If
Next
End Sub
