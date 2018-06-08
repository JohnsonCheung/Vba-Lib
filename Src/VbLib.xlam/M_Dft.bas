Attribute VB_Name = "M_Dft"

Function DftFfn(Ffn0, Optional Ext$ = ".txt", Optional Pth0$, Optional Fdr$)
If Ffn0 <> "" Then DftFfn = Ffn0: Exit Function
Dim Pth$: Pth = DftPth(Pth0)
DftFfn = Pth & TmpNm & Ext
End Function

Function DftPth$(Optional Pth0$, Optional Fdr$)
If Pth0 <> "" Then DftPth = Pth0: Exit Function
Stop
'DftPth = TmpPth(Fdr)
End Function

Function DftNy(Ny0) As String()
If IsStr(Ny0) Then
   DftNy = SslSy(Ny0)
   Exit Function
End If
If IsSy(Ny0) Then
   DftNy = Ny0
End If
End Function

