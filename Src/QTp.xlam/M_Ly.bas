Attribute VB_Name = "M_Ly"
Option Explicit
Property Get LyGpAy(A$(), LinPfx$) As Gp()
Dim J%, O() As Lnx, M() As Lnx
For J = 0 To UB(A)
    Dim Lin$
    Lin = A(J)
    If HasPfx(Lin, LinPfx) Then
        If Sz(M) > 0 Then
            PushObjAy O, M
        End If
        Erase M
    Else
        PushObj M, Lnx(Lin, J)
    End If
Next
If Sz(M) > 0 Then
    PushObjAy O, M
End If
LyGpAy = Gp(O)
End Property
Property Get LyLnxAy(A$()) As Lnx()
Dim O() As Lnx, J%
For J = 0 To UB(A)
    PushObj O, Lnx(A(J), J)
Next
LyLnxAy = O
End Property
