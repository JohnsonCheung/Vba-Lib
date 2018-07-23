Attribute VB_Name = "M_Lo"
Option Explicit

Function LoDrs(A As ListObject) As Drs
Set LoDrs = Drs(LoFny(A), LoDry(A))
End Function

Sub LoBrw(A As ListObject)
Stop
'DrsByLo(A).Brw
End Sub
