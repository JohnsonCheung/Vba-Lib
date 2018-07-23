Attribute VB_Name = "JDta"
Option Explicit

Property Get Drs(Fny0, Dry()) As Drs
Dim O As New Drs
Set Drs = O.Init(Fny0, Dry)
End Property

Property Get Ds(A() As Dt, Optional DsNm$ = "Ds") As Ds
Dim O As New Ds
Set Ds = O.Init(A, DsNm)
End Property

Property Get Dt(DtNm$, Fny0, Dry()) As Dt
Dim O As New Dt
Set Dt = O.Init(DtNm, Fny0, Dry)
End Property
