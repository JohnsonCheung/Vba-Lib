Attribute VB_Name = "M_S1S2Ay"
Option Explicit

Function S1S2AyStr_S1S2Ay(A$) As S1S2()
Dim Ay$(): Ay = Split(A, "|")
Dim O() As S1S2
    Dim I
    For Each I In Ay
        PushObj O, BrkBoth(I, ":")
    Next
S1S2AyStr_S1S2Ay = O
End Function

Function S1S2Ay_Add(A() As S1S2, B() As S1S2) As S1S2()
Dim O() As S1S2
Dim J&
PushObjAy O, A
PushObjAy O, B
S1S2Ay_Add = O
End Function

Function S1S2Ay_Clone(A() As S1S2) As S1S2()
Dim O() As S1S2, I
For Each I In A
    PushObj O, S1S2_Clone(CvS1S2(I))
Next
S1S2Ay_Clone = O
End Function

Function S1S2Ay_Dic(A() As S1S2) As Dictionary
Dim J&, O As New Dictionary
For J = 0 To UB(A)
    With A(J)
        If Not O.Exists(.S1) Then
            O.Add .S1, .S2
        End If
    End With
Next
Set S1S2Ay_Dic = O
End Function

Function S1S2Ay_Sy1(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S1
Next
S1S2Ay_Sy1 = O
End Function

Function S1S2Ay_Sy2(A() As S1S2) As String()
Dim O$(), J&
For J = 0 To UB(A)
   Push O, A(J).S2
Next
S1S2Ay_Sy2 = O
End Function

Function S1S2Ay_SyPair(A() As S1S2) As SyPair
Set S1S2Ay_SyPair = JVb.SyPair(S1S2Ay_Sy1(A), S1S2Ay_Sy2(A))
End Function

Function S1S2Ay_ToStr$(A() As S1S2)
Dim O$(), J%
For J = 0 To UB(A)
    Push O, A(J).ToStr
Next
S1S2Ay_ToStr = Tag("S1S2Ay", JnSpc(O))
End Function

Sub S1S2Ay_Brw(A() As S1S2)
Stop '
'AyBrw S1S2Ay_FmtLy(A)
End Sub

Private Function ZZS1S2Ay() As S1S2()
Dim O() As S1S2
Dim A1$, A2$
Dim I%
I = 0: A1 = "sdklfdlf|lskdfjdf|lskdfj|sldfkj":                 A2 = "sdkdfdfdlfjdf|sldkfjd|l kdf df|   df":          GoSub XX
I = 1: A1 = "sdklfdl df|lskdfjdf|lskdfj|sldfkj":               A2 = "sdklfjsdf|dfdfdf||dfdf|sldkfjd|l kdf df|   df": GoSub XX
I = 2: A1 = "sdsksdlfdf  |df |dfdddf|dflf|lsdf|lskdfj|sldfkj": A2 = "sdklfjdf|sldkfjd|l kdf df|   df": GoSub XX
I = 3: A1 = "sdklfd3lf|lskdfjdf|lskdfj|sldfkj":                A2 = "sdklfjddf||f|sldkfjd|l kdf df|   df": GoSub XX
I = 4: A1 = "sdklfdlf|df|lsk||dfjdf|lskdfj|sldfkj":            A2 = "sdklfjdf|sldkfjdf|d|l kdf df|   df": GoSub XX
ZZS1S2Ay = O
Exit Function
XX:
    PushObj O, S1S2(RplVBar(A1), RplVBar(A2))
    Return

End Function

Private Function ZZS1S2Ay1() As S1S2()
Dim O() As S1S2
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjflsdkjf", "lksdjf")
PushObj O, S1S2("sldjf", "lksdjf")
PushObj O, S1S2("sldjdkjf", "lksdjf")
ZZS1S2Ay1 = O
End Function

Private Sub ZZ_S1S2Ay_FmtLy()
Stop '
'AyBrw S1S2Ay_FmtLy(ZZS1S2Ay)
End Sub

Private Sub ZZ_S1S2Ay_Ly()
Stop '
'AyBrw S1S2Ay_Ly(ZZS1S2Ay1, IsAlignS1:=True)
End Sub
