Attribute VB_Name = "M_Trim"
Option Explicit

Function TrimWhite$(A)
TrimWhite = TrimWhiteL(TrimWhiteL(A))
End Function

Function TrimWhiteL$(A)
Dim J%
    For J = 1 To Len(A)
        If Not IsWhite(Mid(A, J, 1)) Then Exit For
    Next
TrimWhiteL = Left(A, J)
End Function

Function TrimWhiteR$(S)
Dim J%
    Dim A$
    For J = Len(S) To 1 Step -1
        If Not IsWhite(Mid(S, J, 1)) Then Exit For
    Next
    If J = 0 Then Exit Function
TrimWhiteR = Mid(S, J)
End Function
