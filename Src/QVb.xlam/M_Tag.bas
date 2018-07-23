Attribute VB_Name = "M_Tag"
Option Explicit

Function Tag$(TagNm$, S)
If HasPfx(S, TagNm & "(") Then
    If HasSfx(S, ")") Then
        Tag = S
        Exit Function
    End If
End If
End Function

Function Tag_NyStr_ObjAp$(TagNm$, NyStr$, ParamArray ObjAp())
Dim Av(): Av = ObjAp
Tag_NyStr_ObjAp = Tag_Ny_ObjAv(TagNm, SslSy(NyStr), Av)
End Function

Private Function Tag_Ny_ObjAv$(TagNm$, Ny$(), ObjAv())
Ass AyIsSamSz(Ny, ObjAv)
Dim S$
    Dim O$()
    Dim A$, N%
    Dim J%
    For J = 0 To UB(Ny)
        Select Case True
        Case IsNothing(ObjAv(J)): A = "Nothing"
        Case IsEmpty(ObjAv(J)):   A = "Empty"
        Case Else:                A = CallByName(ObjAv(J), "ToStr", VbGet)
        End Select
        Push O, Tag(Ny(J), A)
    Next
    S = JnCrLf(O)
Tag_Ny_ObjAv = Tag(TagNm, S)
End Function
