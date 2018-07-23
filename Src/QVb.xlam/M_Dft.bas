Attribute VB_Name = "M_Dft"
Option Explicit

Function Dft(Val, DftV)
If IsEmp(Val) Then
   Dft = DftV
Else
   Dft = Val
End If
End Function

Function DftFfn$(Ffn0, Optional Ext$ = ".txt", Optional Pth0$, Optional Fdr$)
If Ffn0 <> "" Then DftFfn = Ffn0: Exit Function
Dim Pth$: Pth = DftPth(Pth0)
DftFfn = Pth & TmpNm & Ext
End Function

Function DftNy(Ny0) As String()
If IsStr(Ny0) Then
   DftNy = SslSy(Ny0)
   Exit Function
End If
End Function

Function DftPth$(Optional Pth0$, Optional Fdr$)
If Pth0 <> "" Then DftPth = Pth0: Exit Function
DftPth = TmpPth(Fdr)
End Function

Function DftStr(S$, DftV$)
If S = "" Then
   DftStr = DftV
Else
   DftStr = S
End If
End Function

Function DftTpLy(Tp0) As String()
Stop
'Select Case True
'Case V(Tp0).IsStr: DftTpLy = SplitVBar(Tp0)
'Case IsSy(Tp0):  DftTpLy = Tp0
'Case Else: Stop
'End Select
End Function
