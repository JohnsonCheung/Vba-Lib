Attribute VB_Name = "VbStrHas"
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

Function HasOneOfPfxIgnCas(A, PfxAy) As Boolean
Dim I
For Each I In PfxAy
   If HasPfxIgnCas(A, I) Then HasOneOfPfxIgnCas = True: Exit Function
Next
End Function

Function HasOneOfPfxIgnCas_PfxLvs(A, PfxLvs$) As Boolean
Dim Sy$(): Sy = LvsSy(PfxLvs)
If HasOneOfPfxIgnCas(A, Sy) Then HasOneOfPfxIgnCas_PfxLvs = True: Exit Function
End Function

Function HasPfx(A, Pfx) As Boolean
HasPfx = (Left(A, Len(Pfx)) = Pfx)
End Function

Function HasPfxIgnCas(A, Pfx) As Boolean
HasPfxIgnCas = UCase(Left(A, Len(Pfx))) = UCase(Pfx)
End Function

Function HasSfx(A, Sfx) As Boolean
HasSfx = (Right(A, Len(Sfx)) = Sfx)
End Function
