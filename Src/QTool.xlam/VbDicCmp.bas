Attribute VB_Name = "VbDicCmp"
Option Explicit
Function DCRsltBrw(A As DCRslt)

End Function
Function DCRsltIsSam(A As DCRslt) As Boolean
With A
If .ADif.Count > 0 Then Exit Function
If .BDif.Count > 0 Then Exit Function
If .AExcess.Count > 0 Then Exit Function
If .BExcess.Count > 0 Then Exit Function
End With
DCRsltIsSam = True
End Function
Function DCRsltFmt(A As DCRslt) As String()
With A
Dim A1$(): A1 = DCRsltFmt__AExcess(.AExcess)
Dim A2$(): A2 = DCRsltFmt__BExcess(.BExcess)
Dim A3$(): A3 = DCRsltFmt__Dif(.ADif, .BDif)
Dim A4$(): A4 = DCRsltFmt__Sam(.Sam)
End With
DCRsltFmt = AyAddAp(A1, A2, A3, A4)
End Function
Function DCRsltFmt__AExcess(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, Ly$(), S1$, S2$, S(0) As S1S2
S2 = "!" & "Er AExcess"
For Each K In A.Keys
    S1 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2AyFmt(S)
    PushAy O, Ly
Next
DCRsltFmt__AExcess = O
End Function
Function DCRsltFmt__BExcess(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, Ly$(), S1$, S2$, S(0) As S1S2
S1 = "!" & "Er BExcess"
For Each K In A.Keys
    S2 = K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2AyFmt(S)
    PushAy O, Ly
Next
DCRsltFmt__BExcess = O
End Function
Function DCRsltFmt__Dif(A As Dictionary, B As Dictionary) As String()
If A.Count <> B.Count Then Stop
If A.Count = 0 Then Exit Function
Dim O$(), K, S1$, S2$, S(0) As S1S2, Ly$()
For Each K In A
    S1 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K)
    S2 = "!" & "Er Dif" & vbCrLf & K & vbCrLf & LinesUnderLin(K) & vbCrLf & B(K)
    Set S(0) = S1S2(S1, S2)
    Ly = S1S2AyFmt(S)
    PushAy O, Ly
Next
DCRsltFmt__Dif = O
End Function
Function DCRsltFmt__Sam(A As Dictionary) As String()
If A.Count = 0 Then Exit Function
Dim O$(), K, S() As S1S2
For Each K In A.Keys
    PushObj S, S1S2("*Same", K & vbCrLf & LinesUnderLin(K) & vbCrLf & A(K))
Next
DCRsltFmt__Sam = S1S2AyFmt(S)
End Function
