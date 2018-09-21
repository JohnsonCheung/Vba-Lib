Attribute VB_Name = "M_Vbl"
Option Explicit

Function VblDic(Vbl, Optional JnSep$ = vbCrLf) As Dictionary
Set VblDic = LyDic(SplitVBar(Vbl), JnSep)
End Function

Function VblLines$(Vbl, Optional Pfx$, Optional Ident0%, Optional Sfx$, Optional Wdt0%)
VblLines = JnCrLf(VblLy(VblLines(Vbl), Pfx, Ident0, Sfx, Wdt0))
End Function

Function VblLy(Vbl, Optional Pfx$, Optional Ident0%, Optional Sfx$, Optional Wdt0%) As String()
Ass Vbl_IsVdt(Vbl)
If Vbl = "" Then Exit Function
Dim Wdt%
    Wdt = Vbl_Wdt(Vbl)
    If Wdt < Wdt0 Then
        Wdt = Wdt0
    End If
Dim Ident%
    If Ident < 0 Then
        Ident = 0
    Else
        Ident = Ident0
    End If
    If Pfx <> "" Then
        If Ident < Len(Pfx) Then
            Ident = Len(Pfx) + 1
        End If
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
VblLy = O
End Function

Function Vbl_IsVdt(Vbl) As Boolean
If HasSubStr(Vbl, vbCr) Then Exit Function
If HasSubStr(Vbl, vbLf) Then Exit Function
Vbl_IsVdt = True
End Function

Function Vbl_LasLin$(Vbl)
Vbl_LasLin = AyLasEle(SplitVBar(Vbl))
End Function

Function Vbl_Wdt%(Vbl)
Ass Vbl_IsVdt(Vbl)
Vbl_Wdt = AyWdt(SplitVBar(Vbl))
End Function

Private Sub ZZ_VblLy()
AyDmp VblLy("lksfj|lksdfjldf|lskdlksdflsdf|sdkjf", "Select")
End Sub

Private Sub ZZ_Vbl_Wdt()
Dim Act%: Act = Vbl_Wdt("lksdjf|sldkf|              df")
Ass Act = 16
End Sub

Sub ZZZ__Tst()
ZZ_Vbl_Wdt
End Sub
