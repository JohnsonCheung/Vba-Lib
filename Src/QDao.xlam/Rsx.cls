VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Rsx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private A As Recordset

Property Get Drs() As Drs
Dim O As New Drs
O.Fny = Fny
O.Dry = Dry
Set Drs = O
End Property
