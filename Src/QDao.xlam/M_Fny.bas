Attribute VB_Name = "M_Fny"
Option Explicit
Function FnyQuote(Fny$(), ToQuoteFny$()) As String()
If AyIsEmp(Fny) Then Exit Function
Dim O$(): O = Fny
Dim J%, F
For Each F In O
    If AyHas(ToQuoteFny, F) Then O(J) = Quote(F, "[]")
    J = J + 1
Next
FnyQuote = O
End Function
Function FnyQuoteIfNeed(Fny$()) As String()
If AyIsEmp(Fny) Then Exit Function
Dim O$(), J%, F
O = Fny
For Each F In Fny
    If IsNeedQuote(F) Then O(J) = Quote(CStr(F), "'")
    J = J + 1
Next
FnyQuoteIfNeed = O
End Function
