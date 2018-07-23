Attribute VB_Name = "M_Rmv"
Option Explicit

Function Rmv2Dash$(A)
Rmv2Dash = RTrim(RmvAft(A, "--"))
End Function

Function Rmv3Dash$(A)
Rmv3Dash = RTrim(RmvAft(A, "---"))
End Function

Function RmvAft$(A, Sep$)
RmvAft = Brk1(A, Sep, NoTrim:=True).S1
End Function

Function RmvDblSpc$(A)
Dim O$: O = A
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Function

Function RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Function

Function RmvFstLasChr$(A)
RmvFstLasChr = RmvFstChr(RmvLasChr(A))
End Function

Function RmvFstNChr$(A, Optional N% = 1)
RmvFstNChr = Mid(A, N + 1)
End Function

Function RmvLasChr$(A)
RmvLasChr = RmvLasNChr(A, 1)
End Function

Function RmvLasNChr$(A, N%)
RmvLasNChr = Left(A, Len(A) - 1)
End Function

Function RmvPfx$(S, Pfx)
Dim L%: L = Len(Pfx)
If Left(S, L) = Pfx Then
    RmvPfx = Mid(S, L + 1)
Else
    RmvPfx = S
End If
End Function

Function RmvPfxAy$(A, PfxAy)
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, CStr(Pfx)) Then RmvPfxAy = RmvPfx(A, Pfx): Exit Function
Next
RmvPfxAy = A
End Function

Function RmvSfx$(A, Sfx$)
Dim L%: L = Len(Sfx)
If Right(A, L) = Sfx Then
    RmvSfx = Left(A, Len(A) - L)
Else
    RmvSfx = A
End If
End Function

Sub ZZ__Tst()
ZZ_RmvPfx
End Sub

Private Sub ZZ_RmvPfx()
Ass RmvPfx("aaBB", "aa") = "BB"
End Sub
