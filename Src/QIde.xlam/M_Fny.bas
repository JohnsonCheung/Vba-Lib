Attribute VB_Name = "M_Fny"
Option Explicit
Function FnyOfMthDrs(Optional WithBdyLy As Boolean, Optional WithBdyLines As Boolean) As String()
Dim O$(): O = SplitSpc("MdNm Lno Mdy Ty MthNm")
If WithBdyLy Then Push O, "BdyLy"
If WithBdyLines Then Push O, "BdyLines"
FnyOfMthDrs = O
End Function
