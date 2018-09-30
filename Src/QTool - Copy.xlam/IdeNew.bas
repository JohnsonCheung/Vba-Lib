Attribute VB_Name = "IdeNew"
Option Explicit
Function WhNm(Optional Patn$, Optional Exl$) As WhNm
Dim O As New WhNm
Set O.Re = Re(Patn)
O.ExlAy = SslSy(Exl)
Set WhNm = O
End Function

Function WhMth(Optional WhMdy$, Optional WhKd$, Optional Nm As WhNm) As WhMth
Set WhMth = New WhMth
With WhMth
    .InShtKd = CvWhMthKd(WhKd)
    .InShtMdy = CvWhMdy(WhMdy)
    Set .Nm = Nm
End With
End Function

Function WhMd(WhCmpTy$, Optional Nm As WhNm) As WhMd
Dim O As New WhMd
O.InCmpTy = CvWhCmpTy(WhCmpTy)
Set O.Nm = Nm
Set WhMd = O
End Function

Function WhMdMth_Mth(A As WhMdMth) As WhMth
If Not IsNothing(A) Then Set WhMdMth_Mth = A.Mth
End Function

Function WhMdMth_Md(A As WhMdMth) As WhMd
If Not IsNothing(A) Then Set WhMdMth_Md = A.Md
End Function

Function WhMdMth(Optional Md As WhMd, Optional Mth As WhMth) As WhMdMth
Set WhMdMth = New WhMdMth
With WhMdMth
    Set .Md = Md
    Set .Mth = Mth
End With
End Function

