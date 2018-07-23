Attribute VB_Name = "M_DrsLines"
Option Explicit

Function DrsLines_Drs(A) As Drs
Set DrsLines_Drs = DrsLy_Drs(SplitLines(A))
End Function
