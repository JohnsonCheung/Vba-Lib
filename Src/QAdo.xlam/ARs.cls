VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ARs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Rs As Adodb.Recordset
Friend Property Get Init(A As Record) As ARs
Set Rs = A
Set Init = Me
End Property
Property Get Drs() As Drs
Set Drs = JDta.Drs(Fny, Dry)
End Property

Property Get Dry() As Variant()
Dim O(), F As AFlds
Set F = AFlds(Rs.Fields)
While Not Rs.EOF
    Push O, F.Dr
    Rs.MoveNext
Wend
Dry = O
End Property

Function Fny() As String()
Fny = AFlds(Rs.Fields).Fny
End Function

Private Sub ZZ_Drs()
Dim S$, F As AFb
Set F = AFb(SampleFb_DutyPrepare)
S = "GRANT SELECT ON MSysObjects TO Admin;"
F.Cn.Execute S
Dim Rs As Recordset
Set Rs = F.Cn.Execute("Select * from KE24")
Drs.Brw
End Sub

Friend Sub ZZ__Tst()
ZZ_Drs
End Sub
