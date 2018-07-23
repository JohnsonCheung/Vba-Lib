Attribute VB_Name = "M_Str"
Option Explicit

Function StrAlignL$(S$, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "StrAlignL"
Dim L%: L = Len(S)
If L > W Then
    If ErIfNotEnoughWdt Then
        Stop
        'Er CSub, "Len({S)) > {W}", S, W
    End If
    If DoNotCut Then
        StrAlignL = S
        Exit Function
    End If
End If

If W >= L Then
    StrAlignL = S & Space(W - L)
    Exit Function
End If
If W > 2 Then
    StrAlignL = Left(S, W - 2) + ".."
    Exit Function
End If
StrAlignL = Left(S, W)
End Function

Sub StrBrw(A, Optional Fnn$)
Dim T$: T = TmpFt("StrBrw", Fnn$)
StrWrt A, T
FtBrw T
End Sub

Function StrDup$(N%, S)
Dim O$, J%
For J = 0 To N - 1
    O = O & S
Next
StrDup = O
End Function

Function StrPfx$(A, PfxAy$())
Dim Pfx
For Each Pfx In PfxAy
    If HasPfx(A, CStr(Pfx)) Then StrPfx = Pfx: Exit Function
Next
End Function

Sub StrWrt(A, Ft)
Fso.CreateTextFile(Ft, True).Write A
End Sub
