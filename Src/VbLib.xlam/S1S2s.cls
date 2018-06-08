VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "S1S2s"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A() As S1S2
Property Get U&()
U = UB(A)
End Property
Property Get N&()
N = Sz(A)
End Property
Property Get Init(Ay() As S1S2) As S1S2s
A = Ay
End Property
Property Get ToStr$(Optional Sep$ = " ", Optional S1Wdt%)
Dim O$(), J%
For J = 0 To U
    M_Ay.Push O, A(J).ToStr
Next
ToStr = Tag("S1S2s", JnSpc(O))
End Property

Property Get SyPair() As SyPair
Dim O As New SyPair, Sy1$(), Sy2$(), J&
For J = 0 To U
    With A(J)
        M_Ay.Push Sy1, A(J).S1
        M_Ay.Push Sy2, A(J).S2
    End With
Next
With O
    .Sy1 = Sy1
    .Sy2 = Sy2
End With
Set SyPair = O
End Property
Property Get Wdt1%()
Wdt1 = AyWdt(Sy1)
End Property
Property Get Wdt2%()
Wdt2 = AyWdt(Sy2)
End Property

Property Get Add(A As S1S2s) As S1S2s
'Dim O As S1S2s
'Dim J&
'Set O = Clone
'For J = 0 To UB(A)
'    O.Push A(J)
'Next
'Set Add = O
End Property


Property Get IsEmp() As Boolean
IsEmp = N = 0
End Property

Sub Push(M As S1S2)
Dim NN&
NN = N
ReDim Preserve A(NN)
Set A(N) = M
End Sub

'Function MapStr_Dic(A$) As Dictionary
'Set MapStr_Dic = S1S2AyStr_Dic(A)
'End Function
'
'Function NewS1S2(S1, S2) As S1S2
'NewS1S2.S1 = S1
'NewS1S2.S2 = S2
'End Function
'
'Function NewS1S2Ay(U&) As S1S2()
'If U <= 0 Then Exit Function
'Dim O() As S1S2
'ReDim O(U)
'NewS1S2Ay = O
'End Function
'
'Function NewS1S2Opt(S1, S2, Som As Boolean) As S1S2Opt
'With NewS1S2Opt
'    .Som = Som
'    .S1 = S1
'    .S2 = S2
'End With
'End Function
'
'Function S1S2Ay_Dic(A() As S1S2) As Dictionary
'Dim J&, O As New Dictionary
'For J = 0 To S1S2_UB(A)
'    With A(J)
'        O.Add .S1, .S2
'    End With
'Next
'Set S1S2Ay_Dic = O
'End Function
'Function S1S2AyStr_Dic(A$) As Dictionary
'Set S1S2AyStr_Dic = S1S2Ay_Dic(S1S2AyStr_S1S2Ay(A))
'End Function
'
'Function S1S2AyStr_S1S2Ay(A$) As S1S2()
'Dim Ay$(): Ay = Split(A, "|")
'Dim O() As S1S2
'    Dim I
'    For Each I In Ay
'        S1S2_Push O, BrkBoth(I, ":")
'    Next
'S1S2AyStr_S1S2Ay = O
'End Function
'
'Function S1S2AyStr_SyPair(A$) As SyPair
'S1S2AyStr_SyPair = S1S2Ay_SyPair(S1S2AyStr_S1S2Ay(A))
'End Function
'
'Function S1S2Ay_AddAsLy(A() As S1S2, Optional Sep$ = "") As String()
'Dim O$(), J&
'For J = 0 To S1S2_UB(A)
'   Push O, A(J).S1 & Sep & A(J).S2
'Next
'S1S2Ay_AddAsLy = O
'End Function
'
'Function S1S2Ay_S1LinesWdt%(A() As S1S2)
'S1S2Ay_S1LinesWdt = LinesAy_Wdt(S1S2Ay_Sy1(A))
'End Function
'
'Function S1S2Ay_S2LinesWdt%(A() As S1S2)
'S1S2Ay_S2LinesWdt = LinesAy_Wdt(S1S2Ay_Sy2(A))
'End Function
'
Property Get Sy1() As String()
Dim O$(), J&
For J = 0 To U
   M_Ay.Push O, A(J).S1
Next
Sy1 = O
End Property

Property Get FmtLy() As String()
Dim W1%: W1 = Wdt1
Dim W2%: W2 = Wdt2
Dim H$: H = Hdr(W1, W2)
FmtLy = LinesLinesLy(H, W1, W2)
End Property

Private Function Hdr$(W1%, W2%)
Hdr = "|" + StrDup(W1 + 2, "-") + "|" + StrDup(W2 + 2, "-") + "|"
End Function


Property Get Sy2() As String()
Dim O$(), J&
For J = 0 To U
   M_Ay.Push O, A(J).S2
Next
Sy2 = O
End Property

Function Sy(Optional Sep$ = " ", Optional IsAlignS1 As Boolean) As String()
If IsEmp Then Exit Function
Dim O$(), U&, W%, J%
ReDim O(U)
If IsAlignS1 Then W = Wdt1
For J = 0 To U
    O(J) = A(J).Lin(Sep, W)
Next
Sy = O
End Function

'
'Function SySep_Align(A$(), Sep$) As String()
''Each element of A containc Sep
'If AyIsEmp(A) Then Exit Function
'Dim A1() As S1S2: A1 = SyS1S2Ay(A, Sep)
'SySep_Align = S1S2Ay_Sy(A1, Sep, IsAlignS1:=True)
'End Function

Private Function LinesLinesLy(H$, W1%, W2%) As String()
Dim O$(), I&
M_Ay.Push O, H
For I = 0 To UB(A)
   PushAy O, A(I).Ly(W1, W2)
   M_Ay.Push O, H
Next
LinesLinesLy = O
End Function

Private Sub ZZ_FmtLy()
Dim Act$()
Dim A As New S1S2s
Dim A1$, A2$
Dim I%
I = 0: A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df": GoSub XX
I = 1: A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub XX
I = 2: A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df": GoSub XX
I = 3: A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df": GoSub XX
I = 4: A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df": GoSub XX

Act = A.FmtLy
Stop
'AyBrw Act
Exit Sub
XX:
    A.Push S1S2(RplVBar(A1), RplVBar(A2))
    Return
End Sub


