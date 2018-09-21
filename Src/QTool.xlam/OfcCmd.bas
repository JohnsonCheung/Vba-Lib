Attribute VB_Name = "OfcCmd"
Option Explicit
Sub Z_BldMovMthCmdBar()
BldMovMthCmdBar "VbAy VbAy"
End Sub
Sub BldMovMthCmdBar(ToMdNy0)
Const Nm$ = "MovMth"
If Not VbeHasCmdBar(CurVbe, Nm) Then VbeCrtCmdBar CurVbe, Nm
CmdBarClrAllCtl MovMthCmdBar
End Sub
Sub CmdBarClrAllCtl(A As CommandBar)
Dim I
For Each I In AyNz(CmdBarCtlAy(A))
    CvCtl(I).Delete
Next
End Sub
Function CvCtl(A) As CommandBarControl
Set CvCtl = A
End Function
Function AyNz(A)
If Sz(A) = 0 Then Set AyNz = New Collection: Exit Function
AyNz = A
End Function
Function ItrAyInto(A, OInto)
Dim O, X
O = OInto
Erase O
For Each X In A
    Push O, X
Next
ItrAyInto = O
End Function
Function CmdBarCtlAy(A As CommandBar) As Control()
CmdBarCtlAy = ItrAyInto(A.Controls, CmdBarCtlAy)
End Function
Function MovMthCmdBar() As CommandBar
Set MovMthCmdBar = CurVbe.CommandBars("MovMth")
End Function
Sub CrtCmdBar()
VbeCrtCmdBar CurVbe, "MovMth"
End Sub
Function VbeHasCmdBar(A As Vbe, Nm$) As Boolean
VbeHasCmdBar = ItrHasNm(A.CommandBars, Nm)
End Function
Function CurVbeHasCmdBar(Nm$) As Boolean
CurVbeHasCmdBar = VbeHasCmdBar(CurVbe, Nm)
End Function
Function HasCmdBar(Nm$)
HasCmdBar = CurVbeHasCmdBar(Nm)
End Function
Sub VbeCrtCmdBar(A As Vbe, Nm$)

End Sub
Function VbeCmdBarNy(A As Vbe) As String()
VbeCmdBarNy = ItrNy(A.CommandBars)
End Function
Function CmdBarCtlNy(A As CommandBar) As String()
End Function
Function CmdBarNy() As String()
CmdBarNy = VbeCmdBarNy(CurVbe)
End Function
