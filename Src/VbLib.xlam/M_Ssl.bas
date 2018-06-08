Attribute VB_Name = "M_Ssl"
Property Get SslSy(Ssl) As String()
SslSy = Split(RmvDblSpc(Trim(Ssl)), " ")
End Property
