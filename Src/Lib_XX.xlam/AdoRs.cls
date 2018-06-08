VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AdoRs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Public Rs As AdoDb.Recordset
'Function Drs() As Drs
'Drs.Fny = Fny
'Drs.Dry = Dry
'End Function
'
'Function Dry() As Variant()
'Dim O(), F As New AdoFlds
'Set F.Flds = Rs.Fields
'While Not Rs.EOF
'    Push O, F.Dr
'    Rs.MoveNext
'Wend
'Dry = O
'End Function
'
'Function Fny() As String()
'Dim F As New AdoFlds
'Set F.Flds = Rs.Fields
'Fny = F.Fny
'End Function
'
'Private Sub Dry__Tst()
'Dim S$
'S = "GRANT SELECT ON MSysObjects TO Admin;"
''CurrentProject.Connection.Execute S
'DryBrw AdoP.Rs(Fb(SampleFb_DutyPrepare).Cn.Execute("Select * from KE24")).Dry
'End Sub
'Friend Sub Tst()
'Dry__Tst
'End Sub
