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

Property Get ReMatch(A As RegExp, S) As MatchCollection
Set ReMatch = A.Execute(S)
End Property

Property Get ReRpl$(A As RegExp, S, R$)
ReRpl = A.Replace(S, R)
End Property

Property Get ReTst(A As RegExp, S) As Boolean
ReTst = A.Test(S)
End Property

Sub ZZ_ReMatch()
Dim A As MatchCollection
Dim R  As RegExp: Set R = Re("m[ae]n")
Set A = ReMatch(R, "alskdflfmensdklf")
Stop
End Sub

Sub ZZ_ReRpl()
Dim R As RegExp: Set R = Re("(.+)(m[ae]n)(.+)")
Dim Act$: Act = ReRpl(R, "a men is male", "$1male$3")
Ass Act = "a male is male"
End Sub
