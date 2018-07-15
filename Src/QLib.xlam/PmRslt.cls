VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PmRslt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private C_Pm As Dictionary, C_ErLy$()

Property Get ErLy() As String()
ErLy = C_ErLy
End Property

Property Get Pm() As Dictionary
Set Pm = C_Pm
End Property

Friend Property Get Init(A As Lnx1s) As PmRslt
Dim A1 As LnxsRslt: Set A1 = A.ValidatePfxAndDupT1("%")
Set C_Pm = A1
C_ErLy = A1.ErLy
Set Init = Me
End Property
