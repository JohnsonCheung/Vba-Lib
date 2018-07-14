Attribute VB_Name = "M_Dft"
Option Explicit

Property Get DftFx$(A$)
If A = "" Then
   Dim O$: O = TmpFx
   DftFx = O
Else
   DftFx = A
End If
End Property

Property Get DftWsNmByFxFstWs$(WsNm0, Fx)
Dim O$
Stop
'If WsNm0 = "" Then O = Xls.Fx(Fx).FstWsNm Else O = WsNm0
DftWsNmByFxFstWs = O
End Property


