Attribute VB_Name = "Sql"
Option Explicit
Public Const M_$ = ""


Property Get Tst() As SqlTst
Static Y As New SqlTst
Set Tst = Y
End Property

Function SqpAnd$(Expr$)
If Expr = "" Then Exit Function
SqpAnd = "|    And " & Expr
End Function

Function SqpExprIn$(Expr$, InLis$)
If InLis = "" Then Exit Function
SqpExprIn = FmtQQ("? in (?)", Expr, InLis)
End Function

Function SqpFm$(T)
SqpFm = "|  From " & T
End Function

Function SqpGp$(ExprVblAy$())
Ass IsVdtVblAy(ExprVblAy)
SqpGp = VblAy_AlignAsLines(ExprVblAy, "|  Group By")
End Function

Function SqpInto$(T)
SqpInto = "|  Into " & T
End Function

Function SqpSel$(Fny$(), ExprVblAy$())
SqpSel = VblAy_AlignAsLines(ExprVblAy, "|Select ", SfxAy:=Fny)
End Function

Function SqpSelDis$(Fny, ExprVblAy)
SqpSelDis = "Select Distinct|"
Stop
End Function

Function SqpSelDisFldLvs$(FldLvs$, ExprVblAy$())
Dim Fny$(): Fny = SslSy(FldLvs)
SqpSelDisFldLvs = SqpSelDis(Fny, ExprVblAy)
End Function

Function SqpSelFldLvs$(FldLvs$, ExprVblAy$())
Dim Fny$(): Fny = SslSy(FldLvs)
SqpSelFldLvs = SqpSel(Fny, ExprVblAy)
End Function

Function SqpSet$(FldLvs$, ExprVblAy$())
Const CSub$ = "SqpSet"
Dim Fny$(): Fny = SslSy(FldLvs)
Ass IsVdtVblAy(ExprVblAy)
If Sz(Fny) <> Sz(ExprVblAy) Then Er CSub, "{Sz1}-{FldLvs} <> {FldLvs} and {Sz2}-{ExprVblAy}", Sz(Fny), FldLvs, Sz(ExprVblAy), ExprVblAy
Dim AFny$()
    AFny = AyAlignL(Fny)
    AFny = AyAddSfx(AFny, " = ")
Dim W%
    W = VblAyWdt(ExprVblAy)
Dim Ident%
    W = AyWdt(AFny)
Dim Ay$()
    Dim J%, U%, S$
    U = UB(AFny)
    For J = 0 To U
        If J = U Then
            S = ""
        Else
            S = ","
        End If
        Push Ay, VblAlign(ExprVblAy(J), Pfx:=AFny(J), IdentOpt:=Ident, WdtOpt:=W, Sfx:=S)
    Next
Dim Vbl$
    Dim Ay1$()
    Dim P$
    For J = 0 To U
        If J = 0 Then P = "|  Set" Else P = ""
        Push Ay1, VblAlign(Ay(J), Pfx:=P, IdentOpt:=6)
    Next
    Vbl = JnVBar(Ay1)
SqpSet = Vbl
End Function

Function SqpUpd$(T)
SqpUpd = "Update " & T
End Function

Function SqpWh$(Expr)
SqpWh = "|  Where " & Expr
End Function

Function SqpWhBetStr$(FldNm$, FmStr$, ToStr$)
SqpWhBetStr = FmtQQ("|  Where ? Between '?' and '?'", FldNm, FmStr, ToStr)
End Function

Private Function ZZExprVblAy() As String()
ZZExprVblAy = ApSy("F1-Expr", "F2-Expr   AA|BB    X|DD       Y", "F3-Expr  x")
End Function

Private Function ZZFny() As String()
ZZFny = SplitSpc("F1 F2 F3xxxxx")
End Function

Sub SqpGp__Tst()
Dim ExprVblAy$()
    Push ExprVblAy, "1lskdf|sdlkfjsdfkl sldkjf sldkfj|lskdjf|lskdjfdf"
    Push ExprVblAy, "2dfkl sldkjf sldkdjf|lskdjfdf"
    Push ExprVblAy, "3sldkfjsdf"
AyDmp SplitVBar(SqpGp(ExprVblAy))
End Sub

Private Sub SqpSel__Tst()
Debug.Print RplVBar(SqpSel(ZZFny, ZZExprVblAy))
End Sub

Sub SqpSet__Tst()
Dim Fny$(), ExprVblAy$()
Fny = SslSy("a b c d")
Push ExprVblAy, "1sdfkl|lskdfj|skldfjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "2sdfkl|lskdfjdf| sdf"
Push ExprVblAy, "3sdfkl|fjskldfjs dflkjsdf| sdf"
Push ExprVblAy, "4sf| sdf"

Dim Act$
    Act = SqpSel(Fny, ExprVblAy)
Debug.Print RplVBar(Act)
End Sub
