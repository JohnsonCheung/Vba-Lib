Attribute VB_Name = "M_SampleDb"
Option Explicit
Property Get SampleDb_DutyPrepare() As Database
Set SampleDb_DutyPrepare = DbEng.OpenDatabase(SampleFb_DutyPrepare)
End Property
