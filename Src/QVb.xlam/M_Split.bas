Attribute VB_Name = "M_Split"
Option Explicit

Function SplitComma(A, Optional NoTrim As Boolean) As String()
If NoTrim Then
    SplitComma = Split(A, ",")
Else
    Stop
'    SplitComma = AyTrim(Split(A, ","))
End If
End Function

Function SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Function

Function SplitLf(A) As String()
SplitLf = Split(A, vbLf)
End Function

Function SplitLines(A) As String()
Dim B$: B = Replace(A, vbCrLf, vbLf)
SplitLines = SplitLf(B)
End Function

Function SplitSpc(A) As String()
SplitSpc = Split(A, " ")
End Function

Function SplitVBar(A, Optional Trim As Boolean) As String()
If Trim Then
    SplitVBar = AyTrim(Split(A, "|"))
Else
    SplitVBar = Split(A, "|")
End If
End Function


