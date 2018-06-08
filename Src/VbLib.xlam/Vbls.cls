VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Vbls"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'SpecStr:Vbl: is Vertical-Bar-Lin.  It is a string without VbCr and VbLf.
'SpecStr:Vbl: It uses | as VbCrLf.  It can be converted to Lines.
Option Explicit
Private Vbl$
Property Get Init(A) As Vbl
Vbl = A
Set Init = Me
End Property
Property Get IsVdtVbl(Vbl) As Boolean
If HasSubStr(Vbl, vbCr) Then Exit Property
If HasSubStr(Vbl, vbLf) Then Exit Property
IsVdtVbl = True
End Property

Property Get IsVdtVblAy(VblAy$()) As Boolean
If AyIsEmp(VblAy) Then IsVdtVblAy = True: Exit Property
Dim Vbl
For Each Vbl In VblAy
    If Not IsVdtVbl(Vbl) Then Exit Property
Next
IsVdtVblAy = True
End Property

Property Get VblAlign$(Vbl$, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%)
VblAlign = JnVBar(VblAlignAsLy(Vbl, Pfx, IdentOpt, Sfx, WdtOpt))
End Property

Property Get VblAlignAsLines$(Vbl$, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%)
VblAlignAsLines = JnCrLf(VblAlignAsLy(Vbl, Pfx, IdentOpt, Sfx, WdtOpt))
End Property

Property Get VblAlignAsLy(Vbl$, Optional Pfx$, Optional IdentOpt%, Optional Sfx$, Optional WdtOpt%) As String()
Ass IsVdtVbl(Vbl)
If IsEmp(Vbl) Then Exit Property
Dim Wdt%
    Dim W%
    W = VblWdt(Vbl)
    If W > WdtOpt Then
        Wdt = W
    Else
        Wdt = WdtOpt
    End If
Dim Ident%
    If Ident < 0 Then
        Ident = 0
    Else
        Ident = IdentOpt
    End If
Dim O$()
    Dim Ay$()
    Ay = SplitVBar(Vbl)
    Dim J%, A$, U&, S$, S1$, P$
    U = UB(Ay)
    P = IIf(Pfx <> "", Pfx & " ", "")
    S1 = Space(Ident)
    For J = 0 To U
        If J = 0 Then
            S = AlignL(P, Ident, DoNotCut:=True)
        Else
            S = S1
        End If
        A = S & AlignL(Ay(J), Wdt, ErIfNotEnoughWdt:=True)
        If J = U Then
            A = A & " " & Sfx
        End If
        Push O, A
    Next
VblAlignAsLy = O
End Property

Property Get VblAyWdt%(VblAy$())
Dim W%(), J%
For J = 0 To UB(VblAy)
    Push W, VblWdt(VblAy(J))
Next
VblAyWdt = AyMax(W)
End Property

Property Get VblAy_AlignAsLines$(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAy, Optional Sep$ = ",")
VblAy_AlignAsLines = JnVBar(VblAy_AlignAsLy(ExprVblAy, Pfx, IdentOpt, SfxAy, Sep))
End Property

Property Get VblAy_AlignAsLy(ExprVblAy$(), Optional Pfx$, Optional IdentOpt%, Optional SfxAyOpt, Optional Sep$ = ",") As String()
Dim NoSfxAy As Boolean
Dim SfxWdt%
Dim SfxAy$()
    NoSfxAy = IsEmp(SfxAy)
    If Not NoSfxAy Then
        Ass IsSy(SfxAyOpt)
        SfxAy = AyAlignL(SfxAyOpt)
        Dim U%, J%: U = UB(SfxAy)
        For J = 0 To U
            If J <> U Then
                SfxAy(J) = SfxAy(J) & Sep
            End If
        Next
    End If
Ass IsVdtVblAy(ExprVblAy)
Dim Ident%
    If IdentOpt > 0 Then
        Ident = IdentOpt
    Else
        Ident = 0
    End If
    If Ident = 0 Then
        If Pfx <> "" Then
            Ident = Len(Pfx)
        End If
    End If
Dim O$(), P$, S$
U = UB(ExprVblAy)
Dim W%
    W = VblAyWdt(ExprVblAy)
For J = 0 To U
    If J = 0 Then P = Pfx Else P = ""
    If NoSfxAy Then
        If J = U Then S = "" Else S = Sep
    Else
        If J = U Then S = SfxAy(J) Else S = SfxAy(J) & Sep
    End If
    Push O, VblAlign(ExprVblAy(J), IdentOpt:=Ident, Pfx:=P, WdtOpt:=W, Sfx:=S)
Next
VblAy_AlignAsLy = O
End Property

Property Get VblByLines$(Lines)
If HasSubStr(Lines, "|") Then Stop
VblByLines = Replace(Lines, vbCrLf, "|")
End Property

Property Get VblLasLin$(Vbl)
VblLasLin = AyLasEle(SplitVBar(Vbl))
End Property

Property Get VblLines$(Vbl)
VblLines = Replace(Vbl, "|", vbCrLf)
End Property

Property Get VblLy(Vbl) As String()
VblLy = SplitVBar(Vbl)
End Property

Property Get VblWdt%(Vbl$)
Ass IsVdtVbl(Vbl)
VblWdt = AyWdt(VblLy(Vbl))
End Property

