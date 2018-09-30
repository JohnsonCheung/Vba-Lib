Attribute VB_Name = "VbStrSplit"
Option Explicit
Function SplitComma(A) As String()
SplitComma = Split(A, ",")
End Function
Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function
Function SplitDot(A) As String()
SplitDot = Split(A, ".")
End Function
Function SplitSsl(A) As String()
SplitSsl = Split(RplDblSpc(Trim(A)), " ")
End Function
Function SplitVBar(Vbl$) As String()
SplitVBar = Split(Vbl, "|")
End Function
