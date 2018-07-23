Attribute VB_Name = "M_WinAy"
Option Explicit
Sub WinAy_Keep(A)
Dim W As VBIDE.Window
Dim P&(): P = AyMap_Lng(A, "ObjPointer")
For Each W In CurVbe.Windows
    If Not AyHas(P, ObjPtr(W)) Then
        W.Close
    End If
Next
Dim I
For Each I In A
    CvWin(I).Visible = True
Next
End Sub
