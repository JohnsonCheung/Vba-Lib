Attribute VB_Name = "A__"
Option Explicit
Sub RmkAll()
Dim I, Md As CodeModule
Dim NRmk%, Skip%
For Each I In MbrAy
    Set Md = I
    If Rmk(Md) Then
        NRmk = NRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NRmk"; NRmk
Debug.Print "SKip"; Skip
End Sub

Sub UnRmkAll()
Dim I, Md As CodeModule
Dim NUnRmk%, Skip%
For Each I In MbrAy
    Set Md = I
    If UnRmk(Md) Then
        NUnRmk = NUnRmk + 1
    Else
        Skip = Skip + 1
    End If
Next
Debug.Print "NUnRmk"; NUnRmk
Debug.Print "SKip"; Skip
End Sub

Private Function IsAllRemarked(Md As CodeModule) As Boolean
Dim J%, L$
For J = 1 To Md.CountOfLines
    If Left(Md.Lines(J, 1), 1) <> "'" Then Exit Function
Next
IsAllRemarked = True
End Function
Private Function CurPj() As VBProject
Set CurPj = Application.Vbe.ActiveVBProject
End Function
Private Property Get MbrAy() As CodeModule()
Dim O() As CodeModule, I, Cmp As VBComponent
For Each I In CurPj.VBComponents
    Set Cmp = I
    If Not Cmp.Name = "A__" Then
        PushObj O, Cmp.CodeModule
    End If
Next
MbrAy = O
End Property

Private Sub Push(O, M)
Dim N&
    N = Sz(O)
ReDim Preserve O(N)
O(N) = M
End Sub

Private Sub PushObj(O, M As Object)
Dim N&
    N = Sz(O)
ReDim Preserve O(N)
Set O(N) = M
End Sub

Private Function Rmk(Md As CodeModule) As Boolean
Debug.Print "Rmk " & Md.Parent.Name,
If IsAllRemarked(Md) Then
    Debug.Print " No need"
    Exit Function
End If
Debug.Print "<============= is remarked"
Dim J%
For J = 1 To Md.CountOfLines
    Md.ReplaceLine J, "'" & Md.Lines(J, 1)
Next
Rmk = True
End Function

Private Function Sz&(Ay)
On Error Resume Next
Sz = UBound(Ay) + 1
End Function

Private Function UB&(Ay)
UB = Sz(Ay) - 1
End Function

Private Function UnRmk(Md As CodeModule) As Boolean
Debug.Print "UnRmk " & Md.Parent.Name,
If Not IsAllRemarked(Md) Then
    Debug.Print "No need"
    Exit Function
End If
Debug.Print "<===== is unmarked"
Dim J%, L$
For J = 1 To Md.CountOfLines
    L = Md.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    Md.ReplaceLine J, Mid(L, 2)
Next
UnRmk = True
End Function
