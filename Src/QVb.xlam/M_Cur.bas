Attribute VB_Name = "M_Cur"
Option Explicit
Property Get CurVbe() As VBE
Set CurVbe = Excel.Application.VBE
End Property

Property Get CurMd() As CodeModule
Set CurMd = CurVbe.ActiveCodePane.CodeModule
End Property

