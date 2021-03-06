Attribute VB_Name = "M_Push"
Option Explicit

Sub Push(O, M)
Dim N&
    N = Sz(O)
ReDim Preserve O(N)
If IsObject(M) Then
    Set O(N) = M
Else
    O(N) = M
End If
End Sub

Sub PushAp(O, ParamArray Ap())
Dim Av(), I: Av = Ap
For Each I In Av
    Push O, I
Next
End Sub

Sub PushAy(OAy, Ay)
If AyIsEmp(Ay) Then Exit Sub
Dim I
For Each I In Ay
    Push OAy, I
Next
End Sub

Sub PushNoDup(O, M)
If Not AyHas(O, M) Then Push O, M
End Sub

Sub PushNoDupAy(O, Ay)
Dim I
If AyIsEmp(Ay) Then Exit Sub
For Each I In Ay
    PushNoDup O, I
Next
End Sub

Sub PushNonEmp(O, M)
If IsEmp(M) Then Exit Sub
Push O, M
End Sub

Sub PushObj(O, P)
Dim N&: N = Sz(O)
ReDim Preserve O(N)
Set O(N) = P
End Sub

Sub PushObjAy(O, Ay)
Dim J&
For J = 0 To UB(Ay)
    PushObj O, Ay(J)
Next
End Sub

Sub PushOy(O, Oy)
If AyIsEmp(Oy) Then Exit Sub
Dim M
For Each M In Oy
    PushObj O, M
Next
End Sub
