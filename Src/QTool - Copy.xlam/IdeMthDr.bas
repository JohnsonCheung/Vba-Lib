Attribute VB_Name = "IdeMthDr"
Option Explicit
Function CurVbeSrc() As String()
CurVbeSrc = VbeSrc(CurVbe)
End Function
Function VbeSrc(A As Vbe) As String()
Dim P
For Each P In A.VBProjects
    PushAy VbeSrc, PjSrc(CvPj(P))
Next
End Function
Function PjSrc(A As VBProject) As String()
Dim C As VBComponent
For Each C In A.VBComponents
    PushAy PjSrc, MdSrc(C.CodeModule)
Next
End Function


Sub PushWithSz(O, Ay)
If Not IsArray(Ay) Then Stop
If Sz(Ay) = 0 Then Exit Sub
Push O, Ay
End Sub
Private Sub Z_AyInsAy()
Dim Act, Exp, A(), B(), At&
A = Array(1, 2, 3, 4)
B = Array("X", "Z")
At = 1
Exp = Array(1, "X", "Z", 2, 3, 4)

Act = AyInsAy(A, B, At)
Ass IsEqAy(Act, Exp)
End Sub
Private Sub ZZ_AyReSzAt()
Dim Ay(), At&, Cnt&, Act, Exp
Ay = Array(1, 2, 3)
At = 1
Cnt = 3
Exp = Array(1, Empty, Empty, Empty, 2, 3)
Act = AyReSzAt(Ay, At, Cnt)
Ass IsEqAy(Act, Exp)
End Sub
