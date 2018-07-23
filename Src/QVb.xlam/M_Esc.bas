Attribute VB_Name = "M_Esc"
Option Explicit

Function Esc$(A, Fm$, ToStr$)
Const CSub$ = "Esc"
If InStr(A, ToStr) > 0 Then
    Debug.Print ErMsgLines(CSub, "Warning: escaping a {Str} of {FmStrSub} to {ToSubStr} is found that {Str} contains some {ToSubStr}.  This will make the string chagned after EscUn", A, Fm, ToStr)
End If
Esc = Replace(A, Fm, ToStr)
End Function

Function EscCr$(A)
EscCr = Esc(A, vbCr, "\r")
End Function

Function EscCrLf$(A)
EscCrLf = EscCr(EscLf(A))
End Function

Function EscKey$(A)
EscKey = EscCrLf(EscSpc(EscTab(A)))
End Function

Function EscLf$(A)
EscLf = Esc(A, vbLf, "\n")
End Function

Function EscSpc$(A)
EscSpc = Esc(A, " ", "~")
End Function

Function EscTab$(A)
EscTab = Esc(A, vbTab, "\t")
End Function

Function EscUnCr$(A)
EscUnCr = Replace(A, "\r", vbCr)
End Function

Function EscUnSpc$(A)
EscUnSpc = Replace(A, "~", " ")
End Function

Function EscUnTab(A)
EscUnTab = Replace(A, "\t", "~")
End Function
