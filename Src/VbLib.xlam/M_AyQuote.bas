Attribute VB_Name = "M_AyQuote"
Option Explicit
Function AyQuote(Ay, QuoteStr$) As String()
If AyIsEmp(Ay) Then Exit Function
Dim O$(), U&
    U = UB(Ay)
    ReDim O(U)
    Dim J&
    Dim Q1$, Q2$
    BrkQuote(QuoteStr).Asg Q1, Q2
    For J = 0 To U
        O(J) = Q1 & Ay(J) & Q2
    Next
AyQuote = O
End Function

Property Get AyQuoteSqBkt(Ay) As String()
AyQuoteSqBkt = AyQuote(Ay, "[]")
End Property

Property Get AyQuoteDbl(Ay) As String()
AyQuoteDbl = AyQuote(Ay, """")
End Property

Property Get AyQuoteSng(Ay) As String()
AyQuoteSng = AyQuote(Ay, "'")
End Property

