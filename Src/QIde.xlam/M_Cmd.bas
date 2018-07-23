Attribute VB_Name = "M_Cmd"
Option Explicit
Function CmdBarOfMnu() As CommandBar
Set CmdBarOfMnu = CurVbe.CommandBars("Menu Bar")
End Function
Function CmdBtnOfTileH() As CommandBarButton
Set CmdBtnOfTileH = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Horizontally")
End Function
Function CmdBtnOfTileV() As CommandBarButton
Set CmdBtnOfTileV = CmdPopCap_CmdBtn(CmdPopOfWin, "Tile &Vertically")
End Function
Function CmdPopOfWin() As CommandBarPopup
Set CmdPopOfWin = CmdBarCap_CmdPop(CmdBarOfMnu, "&Window")
End Function
Sub CmdBarOfMnu__Tst()
Debug.Print CmdBarOfMnu.Name
End Sub
Sub CmdPopOfWin__Tst()
Debug.Print CmdPopOfWin.Caption
End Sub
