Attribute VB_Name = "VbStrBrk"
Option Explicit
Function Brk2(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk2 = Brk2__X(A, P, Sep, NoTrim)
End Function
Function Brk2__X(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk2__X = S1S2("", A)
    Else
        Set Brk2__X = S1S2("", Trim(A))
    End If
    Exit Function
End If
Set Brk2__X = BrkAt(A, P, Sep, NoTrim)
End Function
Function Brk(A, Sep, Optional IsNoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
If P = 0 Then Stop
Dim S1$, S2$
    S1 = Left(A, P - 1)
    S2 = Mid(A, P + Len(Sep))
If Not IsNoTrim Then
    S1 = Trim(S1)
    S2 = Trim(S2)
End If
Set Brk = S1S2(S1, S2)
End Function
Function Brk1(A, Sep, Optional NoTrim As Boolean) As S1S2
Dim P&: P = InStr(A, Sep)
Set Brk1 = Brk1__X(A, P, Sep, NoTrim)
End Function
Function Brk1__X(A, P&, Sep, NoTrim As Boolean) As S1S2
If P = 0 Then
    If NoTrim Then
        Set Brk1__X = S1S2(A, "")
    Else
        Set Brk1__X = S1S2(Trim(A), "")
    End If
    Exit Function
End If
Set Brk1__X = BrkAt(A, P, Sep, NoTrim)
End Function
Function BrkAt(A, P&, Sep, NoTrim As Boolean) As S1S2
Dim S1$, S2$
S1 = Left(A, P - 1)
S2 = Mid(A, P + Len(Sep))
If NoTrim Then
    Set BrkAt = S1S2(S1, S2)
Else
    Set BrkAt = S1S2(Trim(S1), Trim(S2))
End If
End Function
Sub Brk2Asg(A, Sep$, O1$, O2$)
Dim P%: P = InStr(A, Sep)
If P = 0 Then
    O1 = ""
    O2 = Trim(A)
Else
    O1 = Trim(Left(A, P - 1))
    O2 = Trim(Mid(A, P + 1))
End If
End Sub
Sub BrkAsg(A, Sep$, O1, O2)
With Brk(A, Sep)
    O1 = .S1
    O2 = .S2
End With
End Sub
