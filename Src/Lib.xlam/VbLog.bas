Attribute VB_Name = "VbLog"
Option Explicit

Sub Log(Msg$, Optional FilNum%)
Dim F%
   F = FilNum
   If FilNum = 0 Then F = LogFilNum
Print #F, NowStr & " " & Msg
If FilNum = 0 Then Close #F
End Sub

Sub LogBrw()
FtBrw LogFt
End Sub

Property Get LogFilNum%()
LogFilNum = FtOpnApp(LogFt)
End Property

Property Get LogFt$()
LogFt = LogPth & "Log.txt"
End Property

Property Get LogPth$()
Dim O$: O = WrkPth & "Log\"
PthEns O
LogPth = O
End Property
