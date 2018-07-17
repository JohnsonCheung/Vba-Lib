Attribute VB_Name = "M_Var"
Option Explicit
Property Get VarCellStr$(V, Optional ShwZer As Boolean)
'CellStr is a string can be displayed in a cell
If IsEmp(V) Then Exit Property
If IsStr(V) Then
    VarCellStr = V
    Exit Property
End If
If IsBool(V) Then
    VarCellStr = IIf(V, "TRUE", "FALSE")
    Exit Property
End If

If IsObject(V) Then
    VarCellStr = "[" & TypeName(V) & "]"
    Exit Property
End If
If ShwZer Then
    If IsNumeric(V) Then
        If V = 0 Then VarCellStr = "0"
        Exit Property
    End If
End If
If IsArray(V) Then
    If AyIsEmp(V) Then Exit Property
    VarCellStr = "Ay" & UB(V) & ":" & V(0)
    Exit Property
End If
If InStr(V, vbCrLf) > 0 Then
    VarCellStr = Brk(V, vbCrLf).S1 & "|.."
    Exit Property
End If
VarCellStr = V
End Property

Property Get VarStr$(A)
If IsPrim(A) Then VarStr = A: Exit Property
If IsNothing(A) Then VarStr = "#Nothing": Exit Property
If IsEmpty(A) Then VarStr = "#Empty": Exit Property
If IsObject(A) Then
    Dim T$
    T = TypeName(A)
    Select Case T
    Case "CodeModule"
        Dim M As CodeModule
        Set M = A
        VarStr = FmtQQ("*Md{?}", M.Parent.Name)
        Exit Property
    End Select
    VarStr = "*" & T
    Exit Property
End If

If IsArray(A) Then
    Dim Ay: Ay = A: ReDim Ay(0)
    T = TypeName(Ay(0))
    VarStr = "*[" & T & "]"
    Exit Property
End If
Stop
End Property

Property Get VarLy(V) As String()
If IsPrim(V) Then
   VarLy = ApSy(V)
ElseIf IsArray(V) Then
   VarLy = AySy(V)
ElseIf IsObject(V) Then
   VarLy = ApSy("*Type: " & TypeName(V))
Else
   Stop
End If
End Property


