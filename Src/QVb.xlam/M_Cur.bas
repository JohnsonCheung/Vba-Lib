Attribute VB_Name = "M_Cur"
Option Explicit

Function CurPj() As VBProject
Set CurPj = CurVbe.ActiveVBProject
End Function

Function CurMd() As CodeModule
Set CurMd = CurVbe.ActiveCodePane.CodeModule
End Function

Function CurVbe() As VBE
Set CurVbe = Excel.Application.VBE
End Function
