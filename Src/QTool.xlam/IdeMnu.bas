Attribute VB_Name = "IdeMnu"
Option Explicit

Function BtnzCompile() As CommandBarButton
Dim O As CommandBarButton
Set O = PopzDbg.CommandBar.Controls(1)
If Not HasPfx(O.Caption, "Compi&le") Then Stop
Set BtnzCompile = O
End Function

Function BtnzTileV() As CommandBarButton
Dim O As CommandBarButton
Set O = PopzWin.CommandBar.Controls(3)
If O.Caption <> "Tile &Vertically" Then Stop
Set BtnzTileV = O
End Function

Function BtnzSav() As CommandBarButton
Dim I As CommandBarControl
For Each I In BarzStd.Controls
    If IsPfx(I.Caption, "&Sav") Then Set BtnzSav = I: Exit Function
Next
Stop
End Function
Sub BtnzCompile_Ass(PjNm$)
If BtnzCompile.Caption <> "Compi&le " & PjNm Then Stop
End Sub
Function BarzStd() As CommandBar
Set BarzStd = CurVbe.CommandBars("Standard")
End Function

Function VbeBarNy(A As Vbe) As String()
VbeBarNy = ItrNy(A.CommandBars)
End Function
Function BarzMnu() As CommandBar
Set BarzMnu = VbeBarzMnu(CurVbe)
End Function
Sub ZZ_BarzMnu()
Dim A As CommandBar
Set A = BarzMnu
Stop
End Sub
Sub ZZ_PopzDbg()
Dim A
Set A = PopzDbg
Stop
End Sub
Function PopzDbg() As CommandBarPopup
Set PopzDbg = BarzMnu.Controls("Debug")
End Function
Function PopzWin() As CommandBarPopup
Set PopzWin = BarzMnu.Controls("Window")
End Function
Function VbeBarzMnu(A As Vbe) As CommandBar
Set VbeBarzMnu = A.CommandBars("Menu Bar")
End Function
Function BarNy() As String()
BarNy = VbeBarNy(CurVbe)
End Function
