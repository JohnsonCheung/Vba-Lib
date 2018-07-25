Attribute VB_Name = "JTool"
Option Explicit

Property Get Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(Fny0, Dry)
End Property

Property Get FTNo(Fmno%, Tono%) As FTNo
Dim O As New FTNo
Set FTNo = O.Init(Fmno, Tono)
End Property
Property Get LCC(Lno%, C1%, C2%) As LCC
Dim O As New LCC
Set LCC = O.Init(Lno, C1, C2)
End Property
Property Get LCCOpt(LCC As LCC) As LCCOpt
Dim O As New LCCOpt
Set LCCOpt = O.Init(LCC)
End Property
Property Get FTIx(Fmix%, Toix%) As FTIx
Dim O As New FTIx
Set FTIx = O.Init(Fmix, Toix)
End Property

Property Get Mth(A As CodeModule, MthNm) As Mth
Dim O As New Mth
Set Mth = O.Init(CvMd(A), MthNm)
End Property

Property Get MthBrk(Nm$, Mdy$, Ty$) As MthBrk
Dim O As New MthBrk
Set MthBrk = O.Init(Nm, Mdy, Ty)
End Property

Property Get MthOpt(A As Mth) As MthOpt
Dim O As New MthOpt
Set MthOpt = O.Init(A)
End Property

Property Get S1S2(S1, S2) As S1S2
Dim O As New S1S2
Set S1S2 = O.Init(S1, S2)
End Property
