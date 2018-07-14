Attribute VB_Name = "G_Vb"
Option Explicit
Property Get ZerFill$(N%, NDig%)
ZerFill = Format(N, StrDup(NDig, 0))
End Property


