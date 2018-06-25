Attribute VB_Name = "M_Sy"
Option Explicit

Property Get SyAddAp(ParamArray Str_or_Sy_Ap()) As String()
Dim Av(): Av = Str_or_Sy_Ap
Dim O$(), I
For Each I In Av
    If IsStr(I) Then
        Push O, I
    ElseIf IsSy(I) Then
        PushAy O, I
    Else
        Stop
    End If
Next
SyAddAp = O
End Property

Property Get SyIsAllEleHasPfx(Ay$(), Pfx$) As Boolean
If AyIsEmp(Ay) Then Exit Property
Dim I
For Each I In Ay
   If Not HasPfx(CStr(I), Pfx) Then Exit Property
Next
SyIsAllEleHasPfx = True
End Property

Property Get SyRmvLasChr(Sy$()) As String()
SyRmvLasChr = AyMap_Sy(Sy, "RmvLasChr")
End Property

Property Get SyTrim(Sy$()) As String()
If AyIsEmp(Sy) Then Exit Property
Dim U&
    U = UB(Sy)
Dim O$()
    Dim J&
    ReDim O(U)
    For J = 0 To U
        O(J) = Trim(Sy(J))
    Next
SyTrim = O
End Property

Sub ZZ__Tst()
ZZ_SyTrim
End Sub

Private Sub ZZ_SyTrim()
AyDmp SyTrim(ApSy(1, 2, 3, "  a"))
End Sub
