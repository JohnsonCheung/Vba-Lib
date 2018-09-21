Attribute VB_Name = "IdeMthSubZ"

Function MdSubZLines$(A As CodeModule)
Dim Ny$(): Ny = Md_FunNy_OfPfx_ZZDash(A)
If Sz(Ny) = 0 Then Exit Function
Ny = AySrt(Ny)
Dim O$()
Dim Pfx$
If A.Parent.Type = vbext_ct_ClassModule Then
    Pfx = "Friend "
End If
Push O, ""
Push O, Pfx & "Sub Z__Tst()"
PushAy O, Ny
Push O, "End Sub"
MdSubZLines = Join(O, vbCrLf)
End Function

