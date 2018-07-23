Attribute VB_Name = "M_Lg"
Option Explicit

Sub Lg(Msg$)
Dim F%
   F = FtOpnApp(LgFt)
Print #F, NowStr & " " & Msg
Close #F
End Sub

Sub LgBrw()
FtBrw LgFt
End Sub

Function LgFt$()
LgFt = LgPth & "Log.txt"
End Function

Function LgPth$()
Dim O$:
O = PgmPth: PthEns O
O = O & "Log\": PthEns O
LgPth = O
End Function
