Attribute VB_Name = "M_Ly"
Option Explicit

Function LyWs(Ly$(), Vis As Boolean) As Worksheet
Dim O As Worksheet: Set O = NewWs(Vis:=Vis)
Dim R As Range
Set R = AyRgV(Ly, WsA1(O))
Set LyWs = O
End Function

