Attribute VB_Name = "VbOy"
Option Explicit

Function OyMap(A, MapMthNm$) As Variant()
Dim Obj, J&, O(), U&
U = UB(A)
ReSz O, U
For J = 0 To U
    Asg Run(MapMthNm, A(J)), O(J)
Next
OyMap = O
End Function
Sub OyEachSub1(A, SubNm$, P)
If AyIsEmp(A) Then Exit Sub
Dim O
For Each O In A
    CallByName O, SubNm, VbMethod, P
Next
End Sub
Sub OyEach(A, MapMthNm$)
Dim Obj, J&, O(), U&
U = UB(A)
ReSz O, U
For J = 0 To U
    Run MapMthNm, A(J)
Next
End Sub

Function OyNy(A) As String()
OyNy = OyPrpSy(A, "Name")
End Function
Function OyDr(Oy) As Variant()
OyDr = OyPrpAy(Oy, "Dr")
End Function
Function OyPrpAy(Oy, PrpNm$, Optional Oup)
Dim O
   If Not IsMissing(Oup) Then
       O = Oup
       Erase O
   Else
       O = EmpAy
   End If
   If AyIsEmp(Oy) Then GoTo X
   Dim I
   For Each I In Oy
       Push O, CallByName(I, PrpNm, VbGet)
   Next
X:
   OyPrpAy = O
End Function

Function OyPrpIntAy(ObjAy, PrpNm$) As Integer()
If AyIsEmp(ObjAy) Then Exit Function
Dim O%(), I
For Each I In ObjAy
   Push O, ObjPrp(I, PrpNm)
Next
OyPrpIntAy = O
End Function

Function OyPrpSy(A, PrpNm$) As String()
OyPrpSy = OyPrpAy(A, PrpNm, EmpSy)
End Function

Function OyWhIxSelIntPrp(Oy, WhIx, PrpNm$) As Integer()
OyWhIxSelIntPrp = OyWhIxSelPrp(Oy, WhIx, PrpNm, EmpIntAy)
End Function

Function OyWhIxSelPrp(Oy, WhIx, PrpNm$, OupAy)
Dim Oy1: Oy1 = AyWhIxAy(Oy, WhIx)  ' Oy1 is subset of Oy
OyWhIxSelPrp = OyPrpAy(Oy1, PrpNm, OupAy)
End Function

Function OyWhIxSelSyPrp(Oy, WhIx, PrpNm$) As String()
OyWhIxSelSyPrp = OyWhIxSelPrp(Oy, WhIx, PrpNm, EmpSy)
End Function

Function OyWhPrp(Oy, PrpNm$, PrpVal)
Dim O
   O = Oy
   Erase O
If Not AyIsEmp(Oy) Then
   Dim I
   For Each I In Oy
       If CallByName(I, PrpNm, VbGet) = PrpVal Then PushObj O, I
   Next
End If
OyWhPrp = O
End Function

Function OyWhPrpEqVal(Oy, PrpNm$, EqVal)
Dim O: O = Oy: Erase O
If Not AyIsEmp(Oy) Then
    Dim I, IsSel As Boolean
    For Each I In Oy
        IsSel = ObjPrp(I, PrpNm) = EqVal
        If IsSel Then
            Push O, I
        End If
    Next
End If
OyWhPrpEqVal = O
End Function

Function OyWhPrpEqValSelPrpInt(Oy, WhPrpNm$, EqVal, SelPrpNm$) As Integer()
Dim Oy1: Oy1 = OyWhPrpEqVal(Oy, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpInt = OyPrpIntAy(Oy1, SelPrpNm)
End Function

Function OyWhPrpEqValSelPrpSy(Oy, WhPrpNm$, EqVal, SelPrpNm$) As String()
Dim Oy1: Oy1 = OyWhPrpEqVal(Oy, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpSy = OyPrpSy(Oy1, SelPrpNm)
End Function

Private Sub OyPrp__Tst()
Dim CdPanAy() As CodePane
CdPanAy = OyPrpAy(CurPjx.MdAy, "CodePane", CdPanAy)
Stop
End Sub

Private Sub Tst()
OyPrp__Tst
End Sub
