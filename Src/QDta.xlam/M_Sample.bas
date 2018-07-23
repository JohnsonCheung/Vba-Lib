Attribute VB_Name = "M_Sample"
Option Explicit

Function SampleDr1() As Variant()
SampleDr1 = Array(1, 2, 3)
End Function

Function SampleDr2() As Variant()
SampleDr2 = Array(2, 3, 4)
End Function

Function SampleDr3() As Variant()
SampleDr3 = Array(3, 4, 5)
End Function

Function SampleDrs() As Drs
Set SampleDrs = Drs("A B C", SampleDry)
End Function

Function SampleDry() As Variant()
SampleDry = Array(SampleDr1, SampleDr2, SampleDr3)
End Function

Function SampleDt() As Dt
Set SampleDt = Dt("Sample", "A B C", SampleDry)
End Function
