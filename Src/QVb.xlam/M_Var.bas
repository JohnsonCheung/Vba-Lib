Attribute VB_Name = "M_Var"
Option Explicit

Function VarCellStr$(V, Optional ShwZer As Boolean)
'CellStr is a string can be displayed in a cell
If IsEmp(V) Then Exit Function
If IsArray(V) Then
    Dim N&: N = Sz(V)
    If N = 0 Then
        VarCellStr = "*[0]"
        Exit Function
    End If
    VarCellStr = "*[" & N & "]" & VarCellStr(V(0))
    Exit Function
End If
If IsObject(V) Then
    VarCellStr = TypeName(V)
    Exit Function
End If
VarCellStr = V
End Function

Function VarLy(V) As String()
If IsPrim(V) Then
   VarLy = ApSy(V)
ElseIf IsArray(V) Then
   VarLy = AySy(V)
ElseIf IsObject(V) Then
   VarLy = ApSy("*Type: " & TypeName(V))
Else
   Stop
End If
End Function

Function VarStr$(A)
If IsPrim(A) Then VarStr = A: Exit Function
If IsNothing(A) Then VarStr = "#Nothing": Exit Function
If IsEmpty(A) Then VarStr = "#Empty": Exit Function
If IsObject(A) Then
    Dim T$
    T = TypeName(A)
    Select Case T
    Case "CodeModule"
        Dim M As CodeModule
        Set M = A
        VarStr = FmtQQ("*Md{?}", M.Parent.Name)
        Exit Function
    End Select
    VarStr = "*" & T
    Exit Function
End If
End Function


