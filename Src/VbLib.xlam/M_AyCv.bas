Attribute VB_Name = "M_AyCv"
Option Explicit
Function AyInto(Ay, OInto)
Ass IsArray(Ay)
Ass IsArray(OInto)
If TypeName(Ay) = TypeName(OInto) Then
    AyInto = Ay
    Exit Function
End If
Dim O: O = OInto: Erase O
Dim I
For Each I In Ay
    Push OInto, I
Next
AyInto = OInto
End Function

Function AyIntAy(Ay) As Integer()
AyIntAy = AyInto(Ay, Emp.IntAy)
End Function

Function AyLngAy(Ay) As Long()
AyLngAy = AyInto(Ay, Emp.LngAy)
End Function

Function AySngAy(Ay) As Single()
AySngAy = AyInto(Ay, Emp.SngAy)
End Function
Function AySy(Ay) As String()
AySy = AyInto(Ay, Emp.Sy)
End Function
Function AyBytAy(Ay) As Byte()
AyBytAy = AyInto(Ay, Emp.BytAy)
End Function

Function AyBoolAy(Ay) As Boolean()
AyBoolAy = AyInto(Ay, Emp.BoolAy)
End Function

