Attribute VB_Name = "VbStrLin"
Option Explicit
Type T1AyRstAy
    T1Ay() As String
    RstAy() As String
End Type
Type T1Rst
    T1 As String
    Rst As String
End Type

Function LinesAy_Wdt%(A$())
Dim O%, J&, M%
For J = 0 To UB(A)
   M = Lines(A(J)).Wdt
   If M > O Then O = M
Next
LinesAy_Wdt = O
End Function

Function LnxAy_Ly(A() As Lnx) As String()
Dim J%, O$()
For J = 0 To LnxUB(A)
    Push O, A(J).Lin
Next
LnxAy_Ly = O
End Function

Sub LnxPush(O() As Lnx, M As Lnx)
Dim N&
    N = LnxSz(O)
ReDim Preserve O(N)
    O(N) = M
End Sub

Function LnxSz%(A() As Lnx)
On Error Resume Next
LnxSz = UBound(A) + 1
End Function

Function LnxUB%(A() As Lnx)
LnxUB = LnxSz(A) - 1
End Function

Function LyT1AyRstAy(A$()) As T1AyRstAy
Dim J&, T1Ay$(), RstAy$()
For J = 0 To UB(A)
    With Lin(A(J)).T1Rslt
        Push T1Ay, .T1
        Push RstAy, .Rst
    End With
Next
With LyT1AyRstAy
    .T1Ay = T1Ay
    .RstAy = RstAy
End With
End Function

Function LyHasMajPfx(A$(), MajPfx$) As Boolean
Dim Cnt%, J%
For J = 0 To UB(A)
    If HasPfx(A(J), MajPfx) Then Cnt = Cnt + 1
Next
LyHasMajPfx = Cnt > (Sz(A) \ 2)
End Function

Function LyLnxAy(A$()) As Lnx()
Dim J&, O() As Lnx
If AyIsEmp(A) Then Exit Function
For J = 0 To UB(A)
    LnxPush O, NewLnx(J, A(J))
Next
LyLnxAy = O
End Function

Function NewLnx(Lx&, Lin$) As Lnx
With NewLnx
    .Lx = Lx
    .Lin = Lin
End With
End Function

Private Sub LinRmvT1__Tst()
Ass Lin("  df dfdf  ").RmvT1 = "dfdf"
End Sub

Private Sub LyBrkFstTermAyRstAy__Tst()
Dim A$()
Push A, "lskdfj sldkfj sldfj sldkfj sldf j"
Push A, "lksj flskdj flsdjk fsldjkf"
Dim Act As T1AyRstAy
Act = LyT1AyRstAy(A)
Stop
End Sub
