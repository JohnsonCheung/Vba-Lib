Attribute VB_Name = "M_Esc"
Option Explicit

Property Get Esc$(A, Fm$, ToStr$)
If InStr(A, "\n") > 0 Then
    Stop
    'Debug.Print ErMsgLines("Esc", "Warning: escaping a {Str} of {FmStrSub} to {ToSubStr} is found that {Str} contains some {ToSubStr}.  This will make the string chagned after UnEsc", A, Fm, ToStr)
End If
Esc = Replace(A, Fm, ToStr)
End Property

Property Get EscCr$(A)
EscCr = Esc(A, vbCr, "\r")
End Property

Property Get EscCrLf$(A)
EscCrLf = EscCr(EscLf(A))
End Property

Property Get EscKey$(A)
EscKey = EscCrLf(EscSpc(EscTab(A)))
End Property

Property Get EscLf$(A)
EscLf = Esc(A, vbLf, "\n")
End Property

Property Get EscSpc$(A)
EscSpc = Esc(A, " ", "~")
End Property

Property Get EscTab$(A)
EscTab = Esc(A, vbTab, "\t")
End Property

Property Get UnEscCr$(A)
UnEscCr = Replace(A, "\r", vbCr)
End Property

Property Get UnEscSpc$(A)
UnEscSpc = Replace(A, "~", " ")
End Property

Property Get UnEscTab(A)
UnEscTab = Replace(A, "\t", "~")
End Property
