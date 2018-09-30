Attribute VB_Name = "IdeBtn"
Option Explicit

Sub ShwNxtStmt()
DbgPop.Visible = True
DoEvents
With NxtStmtBtn
'    If .Enabled Then .Execute
    .Execute
End With
End Sub

Function NxtStmtBtn() As CommandBarButton
Set NxtStmtBtn = DbgPop.Controls("Show Next Statement")
End Function
Function SavBtn() As CommandBarButton
Dim I As CommandBarControl
For Each I In StdBar.Controls
    If IsPfx(I.Caption, "&Sav") Then Set SavBtn = I: Exit Function
Next
Stop
End Function
Sub AssCompileBtn(PjNm$)
If CompileBtn.Caption <> "Compi&le " & PjNm Then Stop
End Sub
Function StdBar() As CommandBar
Set StdBar = CurVbe.CommandBars("Standard")
End Function

Function CompileBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = DbgPop.CommandBar.Controls(1)
If Not HasPfx(O.Caption, "Compi&le") Then Stop
Set CompileBtn = O
End Function

Function TileVBtn() As CommandBarButton
Dim O As CommandBarButton
Set O = WinPop.CommandBar.Controls(3)
If O.Caption <> "Tile &Vertically" Then Stop
Set TileVBtn = O
End Function

Function MnuBar() As CommandBar
Set MnuBar = VbeMnuBar(CurVbe)
End Function
Sub ZZ_MnuBar()
Dim A As CommandBar
Set A = MnuBar
Stop
End Sub
Sub ZZ_DbgPop()
Dim A
Set A = DbgPop
Stop
End Sub
Function DbgPop() As CommandBarPopup
Set DbgPop = MnuBar.Controls("Debug")
End Function
Function WinPop() As CommandBarPopup
Set WinPop = MnuBar.Controls("Window")
End Function
Function BarNy() As String()
BarNy = VbeBarNy(CurVbe)
End Function
Sub Z_BldMovMthBar()
BldMovMthBar "VbAy VbAy"
End Sub
Sub BldMovMthBar(ToMdNy0)
Const Nm$ = "MovMth"
If Not VbeHasBar(CurVbe, Nm) Then VbeCrtBar CurVbe, Nm
BarClrAllCtl MovMthBar
End Sub
Sub BarClrAllCtl(A As CommandBar)
Dim I
For Each I In AyNz(BarCtlAy(A))
    CvCtl(I).Delete
Next
End Sub
Function CvCtl(A) As CommandBarControl
Set CvCtl = A
End Function

Function ItrAyInto(A, OIntoAy)
Dim O: O = OIntoAy: Erase O: ItrAyInto = O
Dim I
For Each I In A
    Push ItrAyInto, I
Next
End Function

Function BarCtlAy(A As CommandBar) As Control()
BarCtlAy = ItrAyInto(A.Controls, BarCtlAy)
End Function
Function MovMthBar() As CommandBar
Set MovMthBar = CurVbe.CommandBars("MovMth")
End Function
Sub CrtMovMthBar()
VbeCrtBar CurVbe, "MovMth"
End Sub
Function CurVbeHasBar(Nm$) As Boolean
CurVbeHasBar = VbeHasBar(CurVbe, Nm)
End Function
Function BarCtlNy(A As CommandBar) As String()
End Function

