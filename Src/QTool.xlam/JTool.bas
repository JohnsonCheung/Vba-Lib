Attribute VB_Name = "JTool"
Option Explicit

Property Get Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(Fny0, Dry)
End Property

Property Get FmToLno(FmLno%, ToLno%) As FmToLno
Dim O As New FmToLno
Set FmToLno = O.Init(FmLno, ToLno)
End Property

Property Get FmToLx(FmLx%, ToLx%) As FmToLx
Dim O As New FmToLx
Set FmToLx = O.Init(FmLx, ToLx)
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
