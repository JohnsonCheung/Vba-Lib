VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Cmd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Property Get Brw() As CmdBrw
Static Y As New CmdBrw
Set Brw = Y
End Property
Property Get Lis() As CmdLis
Static Y As New CmdLis
Set Lis = Y
End Property
Property Get Go() As CmdGo
Static Y As New CmdGo
Set Go = Y
End Property
