Attribute VB_Name = "M_Oy"
Option Explicit

Property Get OyCompoundPrpSy(A, PrpSsl$) As String()
Dim O$(), I
If Sz(A) = 0 Then Exit Property
For Each I In A
    Push O, ObjCompoundPrp(A, PrpSsl)
Next
OyCompoundPrpSy = O
End Property

Property Get OyMap(A, MapMthNm$) As Variant()
OyMap = OyMapInto(A, MapMthNm, EmpAy)
End Property

Property Get OyMapInto(A, MapFunNm$, OIntoAy)
Dim Obj, J&, U&
U = UB(A)
Dim O
O = OIntoAy
ReSz O, U
For J = 0 To U
    Asg Run(MapFunNm, A(J)), O(J)
Next
OyMapInto = O
End Property

Property Get OyNy(A) As String()
OyNy = OyPrpSy(A, "Name")
End Property

Property Get OyPrpAy(A, PrpNm$) As Variant()
If Sz(A) = 0 Then Exit Property
Dim O(), I
For Each I In A
    Push O, CallByName(I, PrpNm, VbGet)
Next
OyPrpAy = O
End Property

Property Get OyPrpIntAy(A, PrpNm$) As Integer()
OyPrpIntAy = OyPrp_Into(A, PrpNm, EmpIntAy)
End Property

Property Get OyPrpSrtedUniqAy(A, PrpNm$) As Variant()
OyPrpSrtedUniqAy = AySrt(AyUniq(OyPrpAy(A, PrpNm)))
End Property

Property Get OyPrpSrtedUniqIntAy(A, PrpNm$) As Integer()
OyPrpSrtedUniqIntAy = AySrt(AyUniq(OyPrpIntAy(A, PrpNm)))
End Property

Property Get OyPrpSrtedUniqSy(A, PrpNm$) As Variant()
OyPrpSrtedUniqSy = AySrt(AyUniq(OyPrpSy(A, PrpNm)))
End Property

Property Get OyPrpSy(A, PrpNm$) As String()
OyPrpSy = OyPrp_Into(A, PrpNm, EmpSy)
End Property

Property Get OyPrp_Into(A, PrpNm$, OIntoAy)
Dim J&
Dim O: O = OIntoAy: Erase O
For J = 0 To UB(A)
    Push O, CallByName(A(J), PrpNm, VbGet)
Next
OyPrp_Into = O
End Property

Property Get OySrt_By_CompoundPrp(A, PrpSsl$)
Dim O: O = A: Erase O
Dim Sy$(): Sy = OyCompoundPrpSy(A, PrpSsl)
Dim Ix&(): Ix = AySrtInToIxAy(Sy)
Dim J&
For J = 0 To UB(Ix)
    PushObj O, A(Ix(J))
Next
OySrt_By_CompoundPrp = O
End Property

Property Get OyToStr$(A)
Dim O$(), I
For Each I In A
    Push O, CallByName(I, "ToStr", VbGet)
Next
OyToStr = JnCrLf(O)
End Property

Property Get OyWhIxAy(A, IxAy)
Dim O: O = A: Erase O
Dim U&: U = UB(IxAy)
Dim J&
ReSz O, U
For J = 0 To U
    Asg A(IxAy(J)), O(J)
Next
OyWhIxAy = O
End Property

Property Get OyWhIxSelIntPrp(A, WhIx, PrpNm$) As Integer()
OyWhIxSelIntPrp = OyWhIxSelPrp(A, WhIx, PrpNm, EmpIntAy)
End Property

Property Get OyWhIxSelPrp(A, WhIx, PrpNm$, OupAy)
Dim Oy1: Oy1 = OyWhIxAy(A, WhIx)  ' Oy1 is subset of Oy
OyWhIxSelPrp = OyPrp_Into(Oy1, PrpNm, OupAy)
End Property

Property Get OyWhIxSelSyPrp(A, WhIx, PrpNm$) As String()
OyWhIxSelSyPrp = OyWhIxSelPrp(A, WhIx, PrpNm, EmpSy)
End Property

Property Get OyWhPrp(A, PrpNm$, PrpEqToVal)
Dim O
   O = A
   Erase O
If Not Sz(A) > 0 Then
   Dim I
   For Each I In A
       If CallByName(I, PrpNm, VbGet) = PrpEqToVal Then PushObj O, I
   Next
End If
OyWhPrp = O
End Property

Property Get OyWhPrpEqVal(A, PrpNm$, EqVal)
Dim O: O = A: Erase O
If Sz(A) > 0 Then
    Dim I, IsSel As Boolean
    For Each I In A
        If ObjPrp(I, PrpNm) = EqVal Then
            PushObj O, I
        End If
    Next
End If
OyWhPrpEqVal = O
End Property

Property Get OyWhPrpEqValSelPrpInt(A, WhPrpNm$, EqVal, SelPrpNm$) As Integer()
Dim Oy1: Oy1 = OyWhPrpEqVal(A, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpInt = OyPrpIntAy(Oy1, SelPrpNm)
End Property

Property Get OyWhPrpEqValSelPrpSy(A, WhPrpNm$, EqVal, SelPrpNm$) As String()
Dim Oy1: Oy1 = OyWhPrpEqVal(A, WhPrpNm, EqVal)
OyWhPrpEqValSelPrpSy = OyPrpSy(Oy1, SelPrpNm)
End Property

Property Get Oy_Cat_AyPrp_AsAy(A, AyPrpNm$)
Dim O, J&, I
If Sz(A) = 0 Then Exit Property
O = CallByName(A(0), AyPrpNm, VbGet)
If Not IsArray(O) Then PmEr ' Given AyPrpNm is not of a array-property
For J = 1 To UB(A)  ' from start Ix=1
    I = CallByName(A(J), AyPrpNm, VbGet)
    If Not IsArray(I) Then DtaEr
    PushAy O, I
Next
Oy_Cat_AyPrp_AsAy = O
End Property

Property Get Oy_Map_ByObjGet(A, Obj, GetMthNm$, OIntoAy)
Dim O: O = OIntoAy
Erase O
Dim ArgAy(0), J%
For J = 0 To UB(A)
    Asg A(J), ArgAy(0)
    Push O, CallByName(Obj, GetMthNm, VbGet, ArgAy)
Next
Oy_Map_ByObjGet = O
End Property

Sub OyDoMth(A, Mth$)
Dim J&
For J = 0 To UB(A)
    CallByName A(J), Mth, VbMethod
Next
End Sub

Sub OyEachSubP1(A, SubNm$, Prm)
If Sz(A) = 0 Then Exit Sub
Dim O
For Each O In A
    CallByName O, SubNm, VbMethod, Prm
Next
End Sub

'=======================================================================================
Sub ZZ__Tst()
ZZ_OyPrpAy
End Sub

Private Sub ZZ_OyPrpAy()
Dim CdPanAy() As CodePane
Stop
'CdPanAy = Oy(CurPjx.MdAy).PrpAy("CodePane", CdPanAy)
Stop
End Sub
