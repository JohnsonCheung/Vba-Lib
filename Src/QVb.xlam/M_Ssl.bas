Attribute VB_Name = "M_Ssl"
Option Explicit

Property Get SslJnComma$(Ssl)
SslJnComma = JnComma(SslSy(Ssl))
End Property

Property Get SslJnQuoteComma$(Ssl)
SslJnQuoteComma = JnComma(AyQuote(SslSy(Ssl), "'"))
End Property

Property Get SslSy(Ssl) As String()
SslSy = Split(RmvDblSpc(Trim(Ssl)), " ")
End Property
