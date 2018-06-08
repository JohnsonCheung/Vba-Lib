Attribute VB_Name = "VbS1S2"
Option Explicit
Type SyPair
    Sy1() As String
    Sy2() As String
End Type

Type V2
    V1 As Variant
    V2 As Variant
End Type
Type V3
    V1 As Variant
    V2 As Variant
    V3 As Variant
End Type
Type S1S2
    S1 As String
    S2 As String
End Type
Type S1S2Opt
    Som As Boolean
    S1 As String
    S2 As String
End Type

Function AyPair_IsEqSz(Ay1, Ay2) As Boolean
AyPair_IsEqSz = Sz(Ay1) = Sz(Ay2)
End Function

Function AyPair_S1S2Ay(Ay1, Ay2) As S1S2()
If AyIsEmp(Ay1) Then Exit Function
Dim U&: U = UB(Ay2)
If U <> UB(Ay1) Then Stop
Dim O() As S1S2
ReDim O(U)
Dim J&
For J = 0 To U
   O(J) = NewS1S2(Ay1(J), Ay2(J))
Next
AyPair_S1S2Ay = O
End Function

Function MapStr_Dic(A$) As Dictionary
Set MapStr_Dic = S1S2AyStr_Dic(A)
End Function

Function NewS1S2(S1, S2) As S1S2
NewS1S2.S1 = S1
NewS1S2.S2 = S2
End Function

Function NewS1S2Ay(U&) As S1S2()
If U <= 0 Then Exit Function
Dim O() As S1S2
ReDim O(U)
NewS1S2Ay = O
End Function

Function NewS1S2Opt(S1, S2, Som As Boolean) As S1S2Opt
With NewS1S2Opt
    .Som = Som
    .S1 = S1
    .S2 = S2
End With
End Function

Function S1S2Ay_Dic(A() As S1S2) As Dictionary
Dim J&, O As New Dictionary
For J = 0 To S1S2_UB(A)
    With A(J)
        O.Add .S1, .S2
    End With
Next
Set S1S2Ay_Dic = O
End Function
Function S1S2AyStr_Dic(A$) As Dictionary
Set S1S2AyStr_Dic = S1S2Ay_Dic(S1S2AyStr_S1S2Ay(A))
End Function

Function S1S2AyStr_S1S2Ay(A$) As S1S2()
Dim Ay$(): Ay = Split(A, "|")
Dim O() As S1S2
    Dim I
    For Each I In Ay
        S1S2_Push O, BrkBoth(I, ":")
    Next
S1S2AyStr_S1S2Ay = O
End Function

Function S1S2AyStr_SyPair(A$) As SyPair
S1S2AyStr_SyPair = S1S2Ay_SyPair(S1S2AyStr_S1S2Ay(A))
End Function

Function S1S2Ay_AddAsLy(A() As S1S2, Optional Sep$ = "") As String()
Dim O$(), J&
For J = 0 To S1S2_UB(A)
   Push O, A(J).S1 & Sep & A(J).S2
Next
S1S2Ay_AddAsLy = O
End Function

Function S1S2Ay_S1LinesWdt%(A() As S1S2)
S1S2Ay_S1LinesWdt = LinesAy_Wdt(S1S2Ay_Sy1(A))
End Function

Function S1S2Ay_S2LinesWdt%(A() As S1S2)
S1S2Ay_S2LinesWdt = LinesAy_Wdt(S1S2Ay_Sy2(A))
End Function

Function S1S2Ay_Sy1(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To S1S2_UB(A)
   Push O, A(J).S1
Next
S1S2Ay_Sy1 = O
End Function

Function S1S2Ay_Sy2(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To S1S2_UB(A)
   Push O, A(J).S2
Next
S1S2Ay_Sy2 = O
End Function

Function S1S2Ay_Sy(A() As S1S2, Optional Sep$ = " ", Optional IsAlignS1 As Boolean) As String()
Dim O$(), U&, W%, J%
U = S1S2_UB(A)
O = NewSy(U)
If IsAlignS1 Then W = S1S2Ay_Wdt1(A)
For J = 0 To U
    O(J) = S1S2_Str(A(J), Sep, W)
Next
S1S2Ay_Sy = O
End Function

Function S1S2Ay_SyPair(A() As S1S2) As SyPair
Dim Sy1$(), Sy2$(), J&
For J = 0 To S1S2_UB(A)
    With A(J)
        Push Sy1, A(J).S1
        Push Sy2, A(J).S2
    End With
Next
With S1S2Ay_SyPair
    .Sy1 = Sy1
    .Sy2 = Sy2
End With
End Function

Function S1S2Ay_Wdt1%(A() As S1S2)
S1S2Ay_Wdt1 = AyWdt(S1S2Ay_Sy1(A))
End Function

Function S1S2_Add(A() As S1S2, B() As S1S2) As S1S2()
Dim O() As S1S2
Dim J&
O = A
For J = 0 To S1S2_UB(B)
    S1S2_Push O, B(J)
Next
S1S2_Add = O
End Function

Function S1S2_Str$(A As S1S2, Optional Sep$ = " ", Optional S1Wdt%)
S1S2_Str = AlignL(A.S1, S1Wdt) & Sep & A.S2
End Function

Function S1S2_Sz&(A() As S1S2)
On Error Resume Next
S1S2_Sz = UBound(A) + 1
End Function

Function S1S2_UB&(A() As S1S2)
S1S2_UB = S1S2_Sz(A) - 1
End Function

Function SomS1S2(S1$, S2$) As S1S2Opt
SomS1S2.S1 = S1
SomS1S2.S2 = S2
SomS1S2.Som = True
End Function

Function SyPair_S1S2Ay(Sy1$(), Sy2$()) As S1S2()
Ass AyPair_IsEqSz(Sy1, Sy2)
If AyIsEmp(Sy1) Then Exit Function
Dim U&, O() As S1S2
ReDim O(U)
Dim J&
For J = 0 To UB(Sy1)
    O(J) = NewS1S2(Sy1(J), Sy2(J))
Next
SyPair_S1S2Ay = O
End Function

Function SyS1S2Ay(A$(), Sep$) As S1S2()
Dim O() As S1S2, J%
Dim U&: U = UB(A)
O = NewS1S2Ay(U)
For J = 0 To U
    With Brk1(A(J), Sep)
        O(J) = NewS1S2(.S1, .S2)
    End With
Next
SyS1S2Ay = O
End Function

Function SySep_Align(A$(), Sep$) As String()
'Each element of A containc Sep
If AyIsEmp(A) Then Exit Function
Dim A1() As S1S2: A1 = SyS1S2Ay(A, Sep)
SySep_Align = S1S2Ay_Sy(A1, Sep, IsAlignS1:=True)
End Function
