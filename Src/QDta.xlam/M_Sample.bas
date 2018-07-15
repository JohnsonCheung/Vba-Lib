Attribute VB_Name = "M_Sample"
Option Explicit

Property Get SampleDr1() As Variant()
SampleDr1 = Array(1, 2, 3)
End Property

Property Get SampleDr2() As Variant()
SampleDr2 = Array(2, 3, 4)
End Property

Property Get SampleDr3() As Variant()
SampleDr3 = Array(3, 4, 5)
End Property

Property Get SampleDrs() As Drs
Set SampleDrs = Drs("A B C", SampleDry)
End Property

Property Get SampleDry() As Variant()
SampleDry = Array(SampleDr1, SampleDr2, SampleDr3)
End Property

Property Get SampleDt() As Dt
Set SampleDt = Dt("Sample", "A B C", SampleDry)
End Property
