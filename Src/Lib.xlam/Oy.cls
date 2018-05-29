VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Oy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private B_Oy ' each element is an object

Sub DoMth(Mth$)
Dim J&
For J = 0 To U
    CallByName B_Oy(J), Mth, VbMethod
Next
End Sub

Function MapByObjGet(Obj, GetMthNm$, OIntoAy)

Dim O: O = OIntoAy
Erase O
Dim ArgAy(0), J%
For J = 0 To U
    Asg Oy(J), ArgAy(0)
    Push O, CallByName(Obj, GetMthNm, VbGet, ArgAy)
Next
MapByObjGet = O
End Function
Property Get N&()
N = Sz(B_Oy)
End Property
Private Function U&()
U = N - 1
End Function
Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property
Function CatAyPrpAsAy(AyPrpNm$)
Dim O, J&, I
If IsEmp Then Exit Function
O = CallByName(B_Oy(0), AyPrpNm, VbGet)
If Not IsArray(O) Then PmEr ' Given AyPrpNm is not of a array-property
For J = 1 To U  ' from start Ix=1
    I = CallByName(B_Oy(J), AyPrpNm, VbGet)
    If Not IsArray(I) Then DtaEr
    PushAy O, I
Next
CatAyPrpAsAy = O
End Function

Function Init(Oy) As Oy
Ass IsArray(Oy)
B_Oy = Oy
Set Init = Me
End Function
Function InitByColl(A As VBA.Collection) As Oy
B_Oy = Coll(A).Ay
Set InitByColl = Me
End Function

Function Sy(PrpNm$) As String()
If IsEmp Then Exit Function
Dim O$(), I
For Each I In B_Oy
    Push O, CallByName(I, PrpNm, VbGet)
Next
Sy = O
End Function
Function ToStr$()
Dim O$(), I
For Each I In B_Oy
    Push O, CallByName(I, "ToStr", VbGet)
Next
ToStr = JnCrLf(O)
End Function
Function SrtedUniqAy(PrpNm$) As Variant()
SrtedUniqAy = AySrt(AyUniq(Ay(PrpNm)))
End Function
Function SrtedUniqIntAy(PrpNm$) As Integer()
SrtedUniqIntAy = AySrt(AyUniq(PrpIntAy(PrpNm)))
End Function
Function SrtedUniqSy(PrpNm$) As Variant()
SrtedUniqSy = AySrt(AyUniq(Sy(PrpNm)))
End Function
Function Ay(PrpNm$) As Variant()
If IsEmp Then Exit Function
Dim O(), I
For Each I In B_Oy
    Push O, CallByName(I, PrpNm, VbGet)
Next
Ay = O
End Function

Function Map(MapMthNm$) As Variant()
Dim Obj, J&, O()
ReSz O, U
For J = 0 To U
    Asg Run(MapMthNm, B_Oy(J)), O(J)
Next
Map = O
End Function
Function MapInto(MapFunNm$, OIntoAy)
Dim Obj, J&
Dim O
O = OIntoAy
ReSz O, U
For J = 0 To U
    Asg Run(MapFunNm, B_Oy(J)), O(J)
Next
MapInto = O
End Function

Sub EachSubP1(SubNm$, Prm)
If IsEmp Then Exit Sub
Dim O
For Each O In B_Oy
    CallByName O, SubNm, VbMethod, Prm
Next
End Sub

Sub EachFun(FunNm$)
If IsEmp Then Exit Sub
Dim I
For Each I In B_Oy
    Run FunNm, I
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
   For Each I In B_Oy
       Push O, CallByName(I, PrpNm, VbGet)
   Next
X:
   PrpAy = O
End Function

Function PrpIntAy(PrpNm$) As Integer()
If IsEmp Then Exit Function
Dim O%(), I
For Each I In B_Oy
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
Dim Oy1: Oy1 = AyWhIxAy(B_Oy, WhIx)  ' Oy1 is subset of Oy
WhIxSelPrp = Oy(Oy1).PrpAy(PrpNm, OupAy)
End Function

Function WhIxSelSyPrp(WhIx, PrpNm$) As String()
WhIxSelSyPrp = WhIxSelPrp(WhIx, PrpNm, EmpSy)
End Function

Function WhPrp(PrpNm$, PrpVal)
Dim O
   O = B_Oy
   Erase O
If Not IsEmp Then
   Dim I
   For Each I In B_Oy
       If CallByName(I, PrpNm, VbGet) = PrpVal Then PushObj O, I
   Next
End If
WhPrp = O
End Function

Function WhPrpEqVal(PrpNm$, EqVal)
Dim O: O = B_Oy: Erase O
If Not IsEmp Then
    Dim I, IsSel As Boolean
    For Each I In B_Oy
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
CdPanAy = Oy(CurPj.MdAy).PrpAy("CodePane", CdPanAy)
Stop
End Sub

Sub Tst()
PrpAy__Tst
End Sub

