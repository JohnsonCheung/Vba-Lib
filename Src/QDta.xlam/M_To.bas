Attribute VB_Name = "M_To"
Option Explicit

Function ToCellStr$(V, ShwZer As Boolean)
'CellStr is a string can be displayed in a cell
If QVb.M_Is.IsEmp(V) Then Exit Function
If IsStr(V) Then
    ToCellStr = V
    Exit Function
End If
If IsBool(V) Then
    ToCellStr = IIf(V, "TRUE", "FALSE")
    Exit Function
End If

If IsObject(V) Then
    ToCellStr = "[" & TypeName(V) & "]"
    Exit Function
End If
If ShwZer Then
    If IsNumeric(V) Then
        If V = 0 Then ToCellStr = "0"
        Exit Function
    End If
End If
If IsArray(V) Then
    If AyIsEmp(V) Then Exit Function
    ToCellStr = "Ay" & UB(V) & ":" & V(0)
    Exit Function
End If
If InStr(V, vbCrLf) > 0 Then
    ToCellStr = Brk(V, vbCrLf).S1 & "|.."
    Exit Function
End If
ToCellStr = V
End Function
