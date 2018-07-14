Attribute VB_Name = "M_S1S2Ay"
Option Explicit

Property Get S1S2AyStr_S1S2Ay(A$) As S1S2()
Dim Ay$(): Ay = Split(A, "|")
Dim O() As S1S2
    Dim I
    For Each I In Ay
        PushObj O, BrkBoth(I, ":")
    Next
S1S2AyStr_S1S2Ay = O
End Property

Property Get S1S2Ay_Add(A() As S1S2, B() As S1S2) As S1S2()
Dim O() As S1S2
Dim J&
PushObjAy O, A
PushObjAy O, B
S1S2Ay_Add = O
End Property

Property Get S1S2Ay_Clone(A() As S1S2) As S1S2()
Dim O() As S1S2, I
For Each I In A
    PushObj O, S1S2_Clone(CvS1S2(I))
Next
S1S2Ay_Clone = O
End Property

Property Get S1S2Ay_Dic(A() As S1S2) As Dictionary
Dim J&, O As New Dictionary
For J = 0 To UB(A)
    With A(J)
        If Not O.Exists(.S1) Then
            O.Add .S1, .S2
        End If
    End With
Next
Set S1S2Ay_Dic = O
End Property

Property Get S1S2Ay_FmtLy(A() As S1S2) As String()
Dim W1%: W1 = ZWdt1(A)
Dim W2%: W2 = ZWdt2(A)
Dim H$: H = ZHdr(W1, W2)
S1S2Ay_FmtLy = ZLinesLinesLy(A, H, W1, W2)
End Property

Property Get S1S2Ay_Ly(A() As S1S2, Optional Sep$ = " ", Optional IsAlignS1 As Boolean) As String()
If Sz(A) = 0 Then Exit Property
Dim O$(), U&, W%, J%
U = UB(A)
ReDim O(U)
If IsAlignS1 Then W = ZWdt1(A)
For J = 0 To U
    O(J) = S1S2_Lin(A(J), Sep, W)
Next
S1S2Ay_Ly = O
End Property

''
Property Get S1S2Ay_Sy1(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   M_Ay.Push O, A(J).S1
Next
S1S2Ay_Sy1 = O
End Property

Property Get S1S2Ay_Sy2(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   M_Ay.Push O, A(J).S2
Next
S1S2Ay_Sy2 = O
End Property

Property Get S1S2Ay_SyPair(A() As S1S2) As SyPair
Set S1S2Ay_SyPair = JVb.SyPair(S1S2Ay_Sy1(A), S1S2Ay_Sy2(A))
End Property

Property Get S1S2Ay_ToStr$(A() As S1S2)
Dim O$(), J%
For J = 0 To UB(A)
    Push O, A(J).ToStr
Next
S1S2Ay_ToStr = Tag("S1S2Ay", JnSpc(O))
End Property


Sub S1S2Ay_Brw(A() As S1S2)
AyBrw S1S2Ay_FmtLy(A)
End Sub

Sub ZZ__Tst()
ZZ_S1S2Ay_FmtLy
ZZ_S1S2Ay_Ly
End Sub

Private Function ZHdr$(W1%, W2%)
ZHdr = "|" + StrDup(W1 + 2, "-") + "|" + StrDup(W2 + 2, "-") + "|"
End Function

Private Property Get ZLinesAy_Wdt%(A$())
If AyIsEmp(A) Then Exit Property
Dim O%, I, M%
For Each I In A
   M = Max(M, LinesWdt(I))
Next
ZLinesAy_Wdt = O
End Property

'Function SySep_Align(A$(), Sep$) As String()
''Each element of A containc Sep
'If AyIsEmp(A) Then Exit Function
'Dim A1() As S1S2: A1 = SyS1S2Ay(A, Sep)
'SySep_Align = S1S2Ay_Sy(A1, Sep, IsAlignS1:=True)
'End Function
Private Property Get ZLinesLinesLy(A() As S1S2, H$, W1%, W2%) As String()
Dim O$(), I&
M_Ay.Push O, H
For I = 0 To UB(A)
   PushAy O, A(I).Ly(W1, W2)
   Push O, H
Next
ZLinesLinesLy = O
End Property

Private Property Get ZLinesWdt1%(A() As S1S2)
ZLinesWdt1 = ZLinesAy_Wdt(S1S2Ay_Sy1(A))
End Property

Private Property Get ZLinesWdt2%(A() As S1S2)
ZLinesWdt2 = ZLinesAy_Wdt(S1S2Ay_Sy2(A))
End Property

Private Property Get ZWdt1%(A() As S1S2)
ZWdt1 = AyWdt(S1S2Ay_Sy1(A))
End Property

Private Property Get ZWdt2%(A() As S1S2)
ZWdt2 = AyWdt(S1S2Ay_Sy2(A))
End Property

Private Property Get ZZS1S2Ay() As S1S2()
Dim O() As S1S2
Dim A1$, A2$
Dim I%
I = 0: A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df":          GoSub XX
I = 1: A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub XX
I = 2: A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df": GoSub XX
I = 3: A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df": GoSub XX
I = 4: A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df": GoSub XX
ZZS1S2Ay = O
Exit Property
XX:
    PushObj O, S1S2(RplVBar(A1), RplVBar(A2))
    Return

End Property

Private Property Get ZZS1S2Ay1() As S1S2()
Dim O() As S1S2
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjf", "lksdjf")
PushObj O, S1S2("sldjdkjf", "lksdjf")
ZZS1S2Ay1 = O
End Property

Private Sub ZZ_S1S2Ay_FmtLy()
AyBrw S1S2Ay_FmtLy(ZZS1S2Ay)
End Sub

Private Sub ZZ_S1S2Ay_Ly()
AyBrw S1S2Ay_Ly(ZZS1S2Ay1, IsAlignS1:=True)
End Sub
