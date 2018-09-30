Attribute VB_Name = "VbStrHas"
Option Explicit
Function HasPfx(S, Pfx) As Boolean
HasPfx = Left(S, Len(Pfx)) = Pfx
End Function
Function HasPfxS(S, Pfx) As Boolean
HasPfxS = Left(S, Len(Pfx) + 1) = Pfx & " "
End Function
Function HasSubStr(A, SubStr$) As Boolean
HasSubStr = InStr(A, SubStr) > 0
End Function
Function HasPfxAy(A, PfxAy) As Boolean
Dim P
For Each P In PfxAy
    If HasPfx(A, P) Then HasPfxAy = True: Exit Function
Next
End Function
Function HasSpc(A) As Boolean
HasSpc = InStr(A, " ") > 0
End Function
Function HasBar(Nm$)
HasBar = CurVbeHasBar(Nm)
End Function
