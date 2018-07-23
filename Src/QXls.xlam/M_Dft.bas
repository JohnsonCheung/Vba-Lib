Attribute VB_Name = "M_Dft"
Option Explicit

Function DftFx$(A$)
If A = "" Then
   Dim O$: O = TmpFx
   DftFx = O
Else
   DftFx = A
End If
End Function

Function DftWsNmByFxFstWs$(WsNm0, Fx)
Dim O$
Stop
'If WsNm0 = "" Then O = Xls.Fx(Fx).FstWsNm Else O = WsNm0
DftWsNmByFxFstWs = O
End Function


