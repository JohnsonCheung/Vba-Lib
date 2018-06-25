Attribute VB_Name = "M_Rpl"
Option Explicit

Property Get RplFstChr$(A, By$)
RplFstChr = By & RmvFstChr(A)
End Property

Property Get RplPfx(A, FmPfx, ToPfx)
RplPfx = ToPfx & RmvPfx(A, FmPfx)
End Property

Property Get RplQ$(A, By$)
RplQ = Replace(A, "?", By)
End Property

Property Get RplVBar$(A)
RplVBar = Replace(A, "|", vbCrLf)
End Property

Sub ZZ__Tst()
ZZ_RplPfx
End Sub

Private Sub ZZ_RplPfx()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub
