Attribute VB_Name = "M_Align"
Option Explicit

Function AlignL$(A, W, Optional ErIfNotEnoughWdt As Boolean, Optional DoNotCut As Boolean)
Const CSub$ = "AlignL"
If ErIfNotEnoughWdt And DoNotCut Then
    Stop
    'Er CSub, "Both {ErIfNotEnoughWdt} and {DontCut} cannot be True", ErIfNotEnoughWdt, DoNotCut
End If
Dim S$: S = VarStr(A)
AlignL = StrAlignL(S, W, ErIfNotEnoughWdt, DoNotCut)
End Function

Function AlignR$(S, W%)
Dim L%: L = Len(S)
If W > L Then
    AlignR = Space(W - L) & S
Else
    AlignR = S
End If
End Function
