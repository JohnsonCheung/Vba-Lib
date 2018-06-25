Attribute VB_Name = "M_VblAy"
Option Explicit

'SpecStr:Vbl: is Vertical-Bar-Lin.  It is a string without VbCr and VbLf.
'SpecStr:Vbl: It uses | as VbCrLf.  It can be converted to Lines.
Property Get VblAy_IsVdt(A$()) As Boolean
If Sz(A) = 0 Then VblAy_IsVdt = True: Exit Property
Dim I
For Each I In A
    If Not Vbl_IsVdt(I) Then Exit Property
Next
VblAy_IsVdt = True
End Property

Property Get VblAy_Lines$(A$(), Optional Pfx$, Optional Ident0%, Optional SfxAy0, Optional Sep$ = ",")
VblAy_Lines = JnVBar(VblAy_Ly(A, Pfx, Ident0, SfxAy0, Sep))
End Property

Property Get VblAy_Ly(A$(), Optional Pfx$, Optional Ident0%, Optional SfxAy0, Optional Sep$ = ",") As String()
Ass VblAy_IsVdt(A)
Dim NoSfxAy As Boolean
Dim SfxWdt%
Dim SfxAy$()
Dim U%
    U = UB(A)
    NoSfxAy = IsEmp(SfxAy)
    If Not NoSfxAy Then
        Ass IsSy(SfxAy0)
        SfxAy = AyAlignL(SfxAy0)
        Dim J%
        For J = 0 To U
            If J <> U Then
                SfxAy(J) = SfxAy(J) & Sep
            End If
        Next
    End If
Dim Ident%
    If Ident0 > 0 Then
        Ident = Ident0
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$
Dim W%
    W = VblAy_Wdt(A)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If NoSfxAy Then
        If J = U Then S = "" Else S = Sep
    Else
        If J = U Then S = SfxAy(J) Else S = SfxAy(J) & Sep
    End If
    Push O, VblLines(A(J), Ident0:=Ident, Pfx:=P, Wdt0:=W, Sfx:=S)
Next
VblAy_Ly = O
End Property

Property Get VblAy_Wdt%(A$())
If Sz(A) = 0 Then Exit Property
Dim W%, I
For Each I In A
    W = Max(W, Vbl_Wdt(I))
Next
VblAy_Wdt = W
End Property

Sub ZZ__Tst()
ZZ_VblAy_Wdt
ZZ_VblAy_Ly
End Sub

Private Function ZZAy() As String()
Dim O$()
Push O, "lksdfj|slkdfjsdf|lksdfj"
Push O, "lksj|slkdfjsdf|lksdfj"
Push O, "lkss|slkdfjsdf|lksdfj"
Push O, "lksdfj|slkdfjsdf|lksdfj"
Push O, "lksdfj|slkdfjsdf|lksdfj"
ZZAy = O
End Function

Private Sub ZZ_VblAy_Ly()
Dim A$(): A = ZZAy
AyBrw VblAy_Ly(A, "select")
End Sub

Private Sub ZZ_VblAy_Wdt()
Dim Act%: Act = VblAy_Wdt(ZZAy)
Ass Act = 9
End Sub
