VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Re"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit
'Private A As RegExp
'
'Friend Property Get Init(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As Re
'Set A = New RegExp
'With A
'   .Pattern = Patn
'   .MultiLine = MultiLine
'   .IgnoreCase = IgnoreCase
'   .Global = IsGlobal
'End With
'Set Init = Me
'End Property
'Function Match(S) As MatchCollection
'Set Match = A.Execute(S)
'End Function
'Function Rpl$(S, R$)
'Rpl = A.Replace(S, R)
'End Function
'Function Tst(S) As Boolean
'Tst = A.Test(S)
'End Function
'Sub Tst1()
''ZZ_Match
'ZZ_Rpl
'End Sub
'Private Sub ZZ_Rpl()
'Debug.Print Init("m[ae]n").Rpl("men", "\1 is changed to woman")
'End Sub
'Private Sub ZZ_Match()
'Dim A As MatchCollection
'Set A = Init("(m[ae]n)").Match("alskdflfmensdklf")
'Stop
'End Sub
'
