Attribute VB_Name = "M_Rmv"
Option Explicit

Property Get Rmv2Dash$(A)
Rmv2Dash = RTrim(RmvAft(A, "--"))
End Property

Property Get Rmv3Dash$(A)
Rmv3Dash = RTrim(RmvAft(A, "---"))
End Property

Property Get RmvAft$(A, Sep$)
RmvAft = Brk1(A, Sep, NoTrim:=True).S1
End Property

Property Get RmvDblSpc$(A)
Dim O$: O = A
While HasSubStr(O, "  ")
    O = Replace(O, "  ", " ")
Wend
RmvDblSpc = O
End Property

Property Get RmvFstChr$(A)
RmvFstChr = Mid(A, 2)
End Property

Property Get RmvFstLasChr$(A)
RmvFstLasChr = RmvFstChr(RmvLasChr(A))
End Property

Property Get RmvFstNChr$(A, Optional N% = 1)
RmvFstNChr = Mid(A, N + 1)
End Property

Property Get RmvLasChr$(A)
RmvLasChr = RmvLasNChr(A, 1)
End Property

Property Get RmvLasNChr$(A, N%)
RmvLasNChr = Left(A, Len(A) - 1)
End Property

Property Get RmvPfx$(S, Pfx)
Dim L%: L = Len(Pfx)
If Left(S, L) = Pfx Then
    RmvPfx = Mid(S, L + 1)
Else
    RmvPfx = S
End If
End Property

Property Get RmvPfxAy$(A, PfxAy)
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, CStr(Pfx)) Then RmvPfxAy = RmvPfx(A, Pfx): Exit Property
Next
RmvPfxAy = A
End Property

Property Get RmvSfx$(A, Sfx$)
Dim L%: L = Len(Sfx)
If Right(A, L) = Sfx Then
    RmvSfx = Left(A, Len(A) - L)
Else
    RmvSfx = A
End If
End Property

Sub ZZ__Tst()
ZZ_RmvPfx
End Sub

Private Sub ZZ_RmvPfx()
Ass RmvPfx("aaBB", "aa") = "BB"
End Sub
