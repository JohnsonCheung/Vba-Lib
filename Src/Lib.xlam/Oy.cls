VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Oy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private A
Sub DoMth(Mth$)
If Sz(A) = 0 Then Add = A: Exit Sub
For J = 0 To UB(A)
    CallByName A(J), Mth, VbMethod
Next
End Sub
Function CatAyPrp(AyPrpNm$)
Dim O()
If Sz(A) = 0 Then Exit Function
O = CallByName(A(0), AyPrpNm, VbGet)
For J = 1 To UB(A)
    PushAy O, A(I)
Next
CatAyPrp = O
End Function
Function Add()
Dim O
If Sz(A) = 0 Then Add = A: Exit Function
Set O = A(0)
For J = 1 To UB(A)
    Set O = O.Add(A(J))
Next
Set Add = O
End Function

Function Init(Oy) As Oy
Ass IsArray(Oy)
A = Oy
Set Init = Me
End Function
Function Sy(PrpNm$) As String()
If IsEmp Then Exit Function
Dim O$()
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
Sy = O
End Function
Function SrtedUniqAy(PrpNm$) As Variant()
SrtedUniqAy = AySrt(AyUniq(Ay(PrpNm)))
End Function
Function SrtedUniqIntAy(PrpNm$) As Integer()
SrtedUniqIntAy = AySrt(AyUniq(PrpIntAy(PrpNm)))
End Function
Function SrtedUniqSy(PrpNm$) As Variant()
SrtedUniqIntSy = AySrt(AyUniq(Sy(PrpNm)))
End Function
Function Ay(PrpNm$) As Variant()
If IsEmp Then Exit Function
Dim O()
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
Ay = O
End Function


Function Map(MapMthNm$) As Variant()
Dim Obj, J&, O(), U&
U = UB(A)
ReSz O, U
For J = 0 To U
    Asg Run(MapMthNm, A(J)), O(J)
Next
Map = O
End Function
Sub EachSubP1(SubNm$, Prm)
If AyIsEmp(A) Then Exit Sub
Dim O
For Each O In A
    CallByName O, SubNm, VbMethod, Prm
Next
End Sub
Sub EachMth(MthNm$)
Dim Obj, J&, O(), U&
U = UB(A)
ReSz O, U
For J = 0 To U
    Run MthNm, A(J)
Next
End Sub

Function Ny() As String()
Ny = PrpSy("Name")
End Function
Function Dry() As Variant()
Dry = PrpAy("Dr")
End Function
Function PrpAy(PrpNm$, Optional Oup)
Dim O
   If Not IsMissing(Oup) Then
       O = Oup
       Erase O
   Else
       O = EmpAy
   End If
   If IsEmp Then GoTo X
   Dim I
   For Each I In A
       Push O, CallByName(I, PrpNm, VbGet)
   Next
X:
   PrpAy = O
End Function

Function PrpIntAy(PrpNm$) As Integer()
If IsEmp Then Exit Function
Dim O%(), I
For Each I In A
   Push O, ObjPrp(I, PrpNm)
Next
PrpIntAy = O
End Function

Function PrpSy(PrpNm$) As String()
PrpSy = PrpAy(PrpNm, EmpSy)
End Function

Function WhIxSelIntPrp(WhIx, PrpNm$) As Integer()
WhIxSelIntPrp = WhIxSelPrp(WhIx, PrpNm, EmpIntAy)
End Function

Function WhIxSelPrp(WhIx, PrpNm$, OupAy)
Dim Oy1: Oy1 = AyWhIxAy(A, WhIx)  ' Oy1 is subset of Oy
WhIxSelPrp = Oy(Oy1).PrpAy(PrpNm, OupAy)
End Function

Function WhIxSelSyPrp(WhIx, PrpNm$) As String()
WhIxSelSyPrp = WhIxSelPrp(WhIx, PrpNm, EmpSy)
End Function

Function WhPrp(PrpNm$, PrpVal)
Dim O
   O = A
   Erase O
If Not IsEmp Then
   Dim I
   For Each I In A
       If CallByName(I, PrpNm, VbGet) = PrpVal Then PushObj O, I
   Next
End If
WhPrp = O
End Function

Function WhPrpEqVal(PrpNm$, EqVal)
Dim O: O = A: Erase O
If Not IsEmp Then
    Dim I, IsSel As Boolean
    For Each I In A
        IsSel = ObjPrp(I, PrpNm) = EqVal
        If IsSel Then
            Push O, I
        End If
    Next
End If
WhPrpEqVal = O
End Function

Function WhPrpEqValSelPrpInt(WhPrpNm$, EqVal, SelPrpNm$) As Integer()
Dim Oy1: Oy1 = WhPrpEqVal(WhPrpNm, EqVal)
WhPrpEqValSelPrpInt = Oy(Oy1).PrpIntAy(SelPrpNm)
End Function

Function WhPrpEqValSelPrpSy(WhPrpNm$, EqVal, SelPrpNm$) As String()
Dim Oy1: Oy1 = WhPrpEqVal(WhPrpNm, EqVal)
WhPrpEqValSelPrpSy = Oy(Oy1).PrpSy(SelPrpNm)
End Function

Private Sub PrpAy__Tst()
Dim CdPanAy() As CodePane
CdPanAy = Oy(CurPjx.MdAy).PrpAy("CodePane", CdPanAy)
Stop
End Sub

Sub Tst()
PrpAy__Tst
End Sub

