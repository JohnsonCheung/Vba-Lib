Attribute VB_Name = "VbFsPth"
Option Explicit
Function PthFfnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
PthFfnAy = AyAddPfx(PthFnAy(A, Spec, Atr), A)
End Function

Function PthFnAy(A, Optional Spec$ = "*.*", Optional Atr As FileAttribute) As String()
If Not PthIsExist(A) Then
    Debug.Print FmtQQ("PthFnAy: Given Path(?) does not exit", A)
    Exit Function
End If
Dim O$()
Dim M$
M = Dir(A & Spec)
If Atr = 0 Then
    While M <> ""
       Push O, M
       M = Dir
    Wend
    PthFnAy = O
End If
Ass PthHasPthSfx(A)
While M <> ""
    If GetAttr(A & M) And Atr Then
        Push O, M
    End If
    M = Dir
Wend
PthFnAy = O
End Function

Function PthHasPthSfx(A) As Boolean
PthHasPthSfx = LasChr(A) = "\"
End Function
Function PthIsExist(A) As Boolean
Ass PthHasPthSfx(A)
PthIsExist = Fso.FolderExists(A)
End Function
Sub PthBrw(P)
Shell "Explorer """ & P & """", vbMaximizedFocus
End Sub
Sub PthClrFil(A)
Dim F
For Each F In AyNz(PthFfnAy(A))
   FfnDlt F
Next
End Sub
Sub PthEns(P$)
If Fso.FolderExists(P) Then Exit Sub
MkDir P
End Sub
