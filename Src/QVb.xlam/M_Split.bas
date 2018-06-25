Attribute VB_Name = "M_Split"
Option Explicit

Property Get SplitComma(A, Optional NoTrim As Boolean) As String()
If NoTrim Then
    SplitComma = Split(A, ",")
Else
    Stop
'    SplitComma = AyTrim(Split(A, ","))
End If
End Property

Property Get SplitCrLf(A) As String()
SplitCrLf = Split(A, vbCrLf)
End Property

Property Get SplitLf(A) As String()
SplitLf = Split(A, vbLf)
End Property

Property Get SplitLines(A) As String()
Dim B$: B = Replace(A, vbCrLf, vbLf)
SplitLines = SplitLf(B)
End Property
