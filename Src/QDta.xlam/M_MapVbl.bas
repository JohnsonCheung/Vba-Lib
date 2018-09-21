Attribute VB_Name = "M_MapVbl"
Option Explicit
Function MapVbl_Dic(A) As Dictionary
Dim Ay$(): Ay = SplitVBar(A)
Dim O As New Dictionary
Dim I
If Sz(Ay) > 0 Then
    For Each I In Ay
        With BrkBoth(I, ":")
            O.Add .S1, .S2
        End With
    Next
End If
Set MapVbl_Dic = O
End Function
