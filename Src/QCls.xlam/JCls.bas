Attribute VB_Name = "JCls"
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
Property Get LCCOpt(Som As Boolean, Optional LCC As LCC) As LCCOpt
Dim O As New LCCOpt
Set LCCOpt = O.Init(Som, LCC)
End Property
Property Get FTIx(Fmix%, Toix%) As FTIx
Dim O As New FTIx
Set FTIx = O.Init(Fmix, Toix)
End Property

Property Get MthCpyPrm(A As Mth, ToMd As CodeModule) As MthCpyPrm
Dim O As New MthCpyPrm
Set MthCpyPrm = O.Init(A, ToMd)
End Property
Property Get Mth(A As CodeModule, MthNm) As Mth
Dim O As New Mth
Set Mth = O.Init(A, MthNm)
End Property

Property Get S1S2(S1, S2) As S1S2
Dim O As New S1S2
Set S1S2 = O.Init(S1, S2)
End Property

Property Get SomLCC(LCC As LCC) As LCCOpt
Set SomLCC = LCCOpt(True, LCC)
End Property
Property Get NonLCC() As LCCOpt
Static X As Boolean, Y As LCCOpt
If Not X Then
    X = True
    Set Y = LCCOpt(False)
End If
Set NonLCC = Y
End Property

Property Get Ds(A() As Dt, Optional DsNm$ = "Ds") As Ds
Dim O As New Ds
Set Ds = O.Init(A, DsNm)
End Property

Property Get Dt(DtNm, Fny0, Dry()) As Dt
Dim O As New Dt
Set Dt = O.Init(DtNm, Fny0, Dry)
End Property

