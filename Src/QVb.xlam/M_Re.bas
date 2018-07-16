Attribute VB_Name = "M_Re"
Option Explicit

Property Get Re(Patn$, Optional MultiLine As Boolean, Optional IgnoreCase As Boolean, Optional IsGlobal As Boolean) As RegExp
Dim O As New RegExp
With O
   .Pattern = Patn
   .MultiLine = MultiLine
   .IgnoreCase = IgnoreCase
   .Global = IsGlobal
End With
Set Re = O
End Property

Private Sub ZZ_ReMatch()
Dim A As MatchCollection
Dim R  As RegExp: Set R = Re("m[ae]n")
Set A = R.Execute("alskdflfmensdklf")
Stop
End Sub

Private Sub ZZ_ReRpl()
Dim R As RegExp: Set R = Re("(.+)(m[ae]n)(.+)")
Dim Act$: Act = R.Replace("a men is male", "$1male$3")
Ass Act = "a male is male"
End Sub
