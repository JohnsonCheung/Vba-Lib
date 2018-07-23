Attribute VB_Name = "M_Prm"
Option Explicit
Function PrmAyNy(A() As MthPrm) As String()
Dim J%, O$()
For J = 0 To MthPrmUB(A)
    Push O, A(J).Nm
Next
PrmAyNy = O
End Function
Function PrmTyAsTyNm$(A As PrmTy)
With A
    If .TyChr <> "" Then PrmTyAsTyNm = TyChrAsTyStr(.TyChr): Exit Function
    If .TyAsNm = "" Then
        PrmTyAsTyNm = "Variant"
    Else
        PrmTyAsTyNm = .TyAsNm
    End If
End With
End Function
Function PrmTyShtNm$(RetTy As PrmTy)
Dim Ay$
Dim O$
    With RetTy
        If .IsAy Then Ay = "Ay"
        Select Case .TyChr
        Case "!": O = "Sng"
        Case "@": O = "Cur"
        Case "#": O = "Dbl"
        Case "$": O = "Str"
        Case "%": O = "Int"
        Case "^": O = "LngLng"
        Case "&": O = "Lng"
        End Select
        If O = "" Then
            O = .TyAsNm
        End If
        If O = "" Then
            O = "Var"
        End If
    End With
    Select Case O
    Case "String": O = "Str"
    Case "Integer": O = "Int"
    Case "Long": O = "Lng"
    Case "Currency": O = "Cur"
    Case "Single": O = "Sng"
    Case "Double": O = "Dbl"
    Case "LongLong": O = "Lng"
    End Select
    O = O & Ay
    If O = "StrAy" Then O = "Sy"
PrmTyShtNm = O
End Function
