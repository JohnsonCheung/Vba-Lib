Attribute VB_Name = "IdeMthRmk"
Option Explicit
Private Function IsRemarked(Cxt$()) As Boolean
If Sz(Cxt) = 0 Then Exit Function
If Not IsPfx(Cxt(0), "Stop '") Then Exit Function
Dim L
For Each L In Cxt
    If Left(L, 1) <> "'" Then Exit Function
Next
IsRemarked = True
End Function

Sub MthCxtFT_Rmk(A As Mth, Cxt As FTNo)
If IsRemarked(MdFTLy(A.Md, Cxt)) Then Exit Sub
Dim J%, L$
For J = Cxt.Fmno To Cxt.Tono
    L = A.Md.Lines(J, 1)
    A.Md.ReplaceLine J, "'" & L
Next
A.Md.InsertLines Cxt.Fmno, "Stop" & " '"
End Sub

Sub MthCxtFT_UnRmk(A As Mth, Cxt As FTNo)
If Not IsRemarked(MdFTLy(A.Md, Cxt)) Then Exit Sub
Dim J%, L$
If Not IsPfx(A.Md.Lines(Cxt.Fmno, 1), "Stop '") Then Stop
For J = Cxt.Fmno + 1 To Cxt.Tono
    L = A.Md.Lines(J, 1)
    If Left(L, 1) <> "'" Then Stop
    A.Md.ReplaceLine J, Mid(L, 2)
Next
A.Md.DeleteLines Cxt.Fmno, 1
End Sub

