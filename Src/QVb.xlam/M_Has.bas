Attribute VB_Name = "M_Has"
Option Explicit

Function Has(A, SubStr) As Boolean
Has = InStr(A, SubStr) > 0
End Function

Function HasCrLf(A) As Boolean
HasCrLf = Has(A, vbCrLf)
End Function

Function HasOneOfPfx(A, PfxAy) As Boolean
Dim I
For Each I In PfxAy
   If HasPfx(A, I) Then HasOneOfPfx = True: Exit Function
Next
End Function
Function HasPfxAy(A, PfxAyVbl0, Optional Compare As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
Dim I
For Each I In DftPfxAy(PfxAyVbl0)
   If HasPfx(A, I, Compare) Then HasPfxAy = True: Exit Function
Next
End Function

Function HasPfx(A, Pfx, Optional Compare As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
HasPfx = StrComp(Left(A, Len(Pfx)), Pfx, Compare)
End Function

Function HasSfx(A, Sfx, Optional Compare As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
HasSfx = StrComp(Right(A, Len(Sfx)), Sfx, Compare)
End Function

Function HasSubStr(S, SubStr, Optional Compare As VbCompareMethod = VbCompareMethod.vbTextCompare) As Boolean
HasSubStr = InStr(S, SubStr, Compare) > 0
End Function

Function HasSubStrAy(S, SubStrAy) As Boolean
Dim SubStr
For Each SubStr In SubStrAy
    If HasSubStr(S, SubStr) Then HasSubStrAy = True: Exit Function
Next
End Function

Function HasVBar(S) As Boolean
HasVBar = HasSubStr(S, "|")
End Function
