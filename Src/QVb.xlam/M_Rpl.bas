Attribute VB_Name = "M_Rpl"
Option Explicit

Function RplFstChr$(A, By$)
RplFstChr = By & RmvFstChr(A)
End Function

Function RplPfx(A, FmPfx, ToPfx)
RplPfx = ToPfx & RmvPfx(A, FmPfx)
End Function

Function RplQ$(A, By$)
RplQ = Replace(A, "?", By)
End Function

Function RplVBar$(A)
RplVBar = Replace(A, "|", vbCrLf)
End Function

Sub ZZ__Tst()
ZZ_RplPfx
End Sub

Private Sub ZZ_RplPfx()
Ass RplPfx("aaBB", "aa", "xx") = "xxBB"
End Sub
