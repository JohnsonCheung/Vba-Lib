VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public Rs As ADODB.Recordset
Function Drs() As Drs
Drs.Fny = Fny
Drs.Dry = Dry
End Function

Function Dry() As Variant()
Dim O(), F As New AFlds
Set F.Flds = Rs.Fields
While Not Rs.EOF
    Push O, F.Dr
    Rs.MoveNext
Wend
Dry = O
End Function

Function Fny() As String()
Dim F As New AFlds
Set F.Flds = Rs.Fields
Fny = F.Fny
End Function

