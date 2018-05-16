Attribute VB_Name = "VbS1S2_Fun_S1S2Ay_ToFmtLy"
Option Explicit
Enum e_S1S2FmtOpt
    e_LinesLines = 1
    e_KeyLines
End Enum
Private Function Z12_LinesLines(A() As S1S2, H$, W1%, W2%) As String()
Dim O$(), I&
Push O, H
For I = 0 To S1S2_UB(A)
   PushAy O, Z121_Ly(A(I), W1, W2)
   Push O, H
Next
Z12_LinesLines = O
End Function
Function S1S2_Ly(A As S1S2, Optional KeyWdt0%) As String()
With A
    S1S2_Ly = KeyLines_Ly(.S1, .S2, KeyWdt0)
End With
End Function
Private Function Z13_KeyLines(A() As S1S2, W1%) As String()
Dim J&, O$()
For J = 0 To S1S2_UB(A)
    PushAy O, S1S2_Ly(A(J), W1)
Next
Z13_KeyLines = O
End Function

Function S1S2Ay_FmtLy(A() As S1S2, Optional Op As e_S1S2FmtOpt = e_LinesLines) As String()
S1S2Ay_FmtLy = Z1_S1S2Ay_FmtLy(A, Op)
End Function

Private Function Z1_S1S2Ay_FmtLy(A() As S1S2, Optional Op As e_S1S2FmtOpt = e_LinesLines) As String()
Dim W1%: W1 = S1S2Ay_S1LinesWdt(A)
Select Case Op
Case e_LinesLines
    Dim W2%: W2 = S1S2Ay_S2LinesWdt(A)
    Dim H$: H = Z11_Hdr(W1, W2)
    Z1_S1S2Ay_FmtLy = Z12_LinesLines(A, H, W1, W2)
Case e_KeyLines
    Z1_S1S2Ay_FmtLy = Z13_KeyLines(A, W1)
Case Else
    Stop
End Select
End Function
Private Function Z11_Hdr$(W1%, W2%)
Z11_Hdr = "|" + StrDup(W1 + 2, "-") + "|" + StrDup(W2 + 2, "-") + "|"
End Function

Private Function Z121_Ly(A As S1S2, W1%, W2%) As String()
Dim S1$(), S2$()
S1 = SplitCrLf(A.S1)
S2 = SplitCrLf(A.S2)
Dim M%, J%, O$(), Lin$, A1$, A2$, U1%, U2%
    U1 = UB(S1)
    U2 = UB(S2)
    M = Max(U1, U2)
Dim Spc1$, Spc2$
    Spc1 = Space(W1)
    Spc2 = Space(W2)
For J = 0 To M
   If J > U1 Then
       A1 = Spc1
   Else
       A1 = AlignL(S1(J), W1)
   End If
   If J > U2 Then
       A2 = Spc2
   Else
       A2 = AlignL(S2(J), W2)
   End If
   Lin = "| " + A1 + " | " + A2 + " |"
   Push O, Lin
Next
Z121_Ly = O
End Function

Private Sub S1S2Ay_FmtLy__Tst()
Dim Act$()
Dim A() As S1S2
ReDim A(4)
Dim A1$, A2$
Dim I%
I = 0: A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df": GoSub XX
I = 1: A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub XX
I = 2: A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df": GoSub XX
I = 3: A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df": GoSub XX
I = 4: A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df": GoSub XX

Act = S1S2Ay_FmtLy(A)
AyBrw Act
Exit Sub
XX:
    A(I) = NewS1S2(RplVBar(A1), RplVBar(A2))
    Return
End Sub
Sub Tst()
S1S2Ay_FmtLy__Tst
End Sub
