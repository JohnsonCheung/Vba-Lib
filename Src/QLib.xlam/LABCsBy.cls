VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LABCsBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Property Get Init(ABCAy() As ABC, IsVF As Boolean) As LABCs
Dim O As New LABCs
Set Init = O.Init(ABCAy, IsVF)
End Property

Property Get ByLines(ABCLines$, Optional IsVF As Boolean) As LABCs
Dim O As New LABCs
Set ByLines = O.InitByLines(ABCLines, IsVF)
End Property

Property Get ByT1(T1$, Optional IsVF As Boolean) As LABCs
Dim O As New LABCs
Set ByT1 = O.InitByT1(T1, IsVF)
End Property

Property Get ByVBarLines(ABCVBarLines$, Optional IsVF As Boolean) As LABCs
Set ByVBarLines = ByLines(RplVBar(ABCVBarLines), IsVF)
End Property
