Attribute VB_Name = "VbS1S2_Fun_S1S2Ay_FmtLy"
Option Explicit
Function S1S2Ay_FmtLy(A() As S1S2) As String()
Dim W1%: W1 = S1S2Ay_S1LinesWdt(A)
Dim W2%: W2 = S1S2Ay_S2LinesWdt(A)
Dim H$: H = Hdr(W1, W2)
S1S2Ay_FmtLy = S1S2Ay_LinesLinesLy(A, H, W1, W2)
End Function

Private Function Hdr$(W1%, W2%)
Hdr = "|" + StrDup(W1 + 2, "-") + "|" + StrDup(W2 + 2, "-") + "|"
End Function

Private Function S1S2_Ly(A As S1S2, W1%, W2%) As String()
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
S1S2_Ly = O
End Function

Private Function S1S2Ay_LinesLinesLy(A() As S1S2, H$, W1%, W2%) As String()
Dim O$(), I&
Push O, H
For I = 0 To S1S2_UB(A)
   PushAy O, S1S2_Ly(A(I), W1, W2)
   Push O, H
Next
S1S2Ay_LinesLinesLy = O
End Function

Private Sub ZZ_S1S2Ay_FmtLy()
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
