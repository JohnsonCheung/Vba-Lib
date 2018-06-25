Attribute VB_Name = "M_Dft"
Option Explicit

Property Get Dft(Val, DftV)
If IsEmp(Val) Then
   Dft = DftV
Else
   Dft = Val
End If
End Property

Property Get DftFfn$(Ffn0, Optional Ext$ = ".txt", Optional Pth0$, Optional Fdr$)
If Ffn0 <> "" Then DftFfn = Ffn0: Exit Property
Dim Pth$: Pth = DftPth(Pth0)
DftFfn = Pth & TmpNm & Ext
End Property

Property Get DftLy(Ly0) As String()
If IsStr(Ly0) Then
    DftLy = SplitVBar(Ly0)
ElseIf IsSy(Ly0) Then
    DftLy = Ly0
Else
    PmEr
End If
End Property

Property Get DftNy(Ny0) As String()
If IsStr(Ny0) Then
   DftNy = SslSy(Ny0)
   Exit Property
End If
If IsSy(Ny0) Then
   DftNy = Ny0
End If
End Property

Property Get DftPth$(Optional Pth0$, Optional Fdr$)
If Pth0 <> "" Then DftPth = Pth0: Exit Property
DftPth = TmpPth(Fdr)
End Property

Property Get DftStr(S$, DftV$)
If S = "" Then
   DftStr = DftV
Else
   DftStr = S
End If
End Property
