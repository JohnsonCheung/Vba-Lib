Attribute VB_Name = "M_Has"
Option Explicit

Property Get Has(A, SubStr) As Boolean
Has = InStr(A, SubStr) > 0
End Property

Property Get HasCrLf(A) As Boolean
HasCrLf = Has(A, vbCrLf)
End Property

Property Get HasOneOfPfx(A, PfxAy) As Boolean
Dim I
For Each I In PfxAy
   If HasPfx(A, I) Then HasOneOfPfx = True: Exit Property
Next
End Property

Property Get HasOneOfPfxIgnCas(A, PfxAy) As Boolean
Dim I
For Each I In PfxAy
   If HasPfxIgnCas(A, I) Then HasOneOfPfxIgnCas = True: Exit Property
Next
End Property

Property Get HasOneOfPfxIgnCas_PfxSsl(A, PfxSsl$) As Boolean
Stop
Dim Sy$(): 'Sy = SslSy(PfxSsl)
If HasOneOfPfxIgnCas(A, Sy) Then HasOneOfPfxIgnCas_PfxSsl = True: Exit Property
End Property

Property Get HasPfx(A, Pfx) As Boolean
HasPfx = (Left(A, Len(Pfx)) = Pfx)
End Property

Property Get HasPfxIgnCas(A, Pfx) As Boolean
HasPfxIgnCas = UCase(Left(A, Len(Pfx))) = UCase(Pfx)
End Property

Property Get HasSfx(A, Sfx) As Boolean
HasSfx = (Right(A, Len(Sfx)) = Sfx)
End Property
