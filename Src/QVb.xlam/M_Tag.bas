Attribute VB_Name = "M_Tag"
Option Explicit

Property Get Tag$(TagNm$, S)
If HasPfx(S, TagNm & "(") Then
    If HasSfx(S, ")") Then
        Tag = S
        Exit Property
    End If
End If
Stop
If Has(S, vbCrLf) Then
'    Tag = FmtQQ("?(|?|?)", TagNm, S, TagNm)
Else
'    Tag = FmtQQ("?(?)", TagNm, S)
End If
End Property

Property Get Tag_NyStr_ObjAp$(TagNm$, NyStr$, ParamArray ObjAp())
Dim Av(): Av = ObjAp
Stop
'Tag_NyStr_ObjAp = Tag_Ny_ObjAv(TagNm, SslSy(NyStr), Av)
End Property

Private Property Get Tag_Ny_ObjAv$(TagNm$, Ny$(), ObjAv())
Stop
'Ass AyIsSamSz(Ny, ObjAv)
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
Stop
'    S = JnCrLf(O)
Tag_Ny_ObjAv = Tag(TagNm, S)
End Property
