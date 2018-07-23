Attribute VB_Name = "M_Ssl"
Option Explicit

Function SslJnComma$(Ssl)
SslJnComma = JnComma(SslSy(Ssl))
End Function

Function SslJnQuoteComma$(Ssl)
SslJnQuoteComma = JnComma(AyQuote(SslSy(Ssl), "'"))
End Function

Function SslSy(Ssl) As String()
SslSy = Split(RmvDblSpc(Trim(Ssl)), " ")
End Function
