Attribute VB_Name = "G_Cls"
Option Explicit
Function Sz&(A)
On Error Resume Next
Sz = UBound(A) + 1
End Function
Function DftNy(A) As String()
Select Case True
Case VarType(A) = vbString: DftNy = SslSy(A)
Case VarType(A) = vbString + vbArray
    DftNy = A
Case Else
    Stop
End Select
End Function
Function SslSy(A) As String()
Dim B$, Ay$(), O$(), I
B = Trim(A)
Ay = Split(B, " ")
For Each I In Ay
    If I <> " " Then
        Push O, I
    End If
Next
SslSy = O
End Function
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

