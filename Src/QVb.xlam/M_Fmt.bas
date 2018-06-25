Attribute VB_Name = "M_Fmt"
Option Explicit

Property Get FmtMacro$(MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
FmtMacro = FmtMacroAv(MacroStr, Av)
End Property

Property Get FmtMacroAv$(MacroStr$, Av())
Dim Ay$(): Ay = MacroNy(MacroStr)
Dim O$: O = MacroStr
Dim J%, I
For Each I In Ay
    O = Replace(O, I, Av(J))
    J = J + 1
Next
FmtMacroAv = O
End Property

Property Get FmtMacroDic$(MacroStr$, Dic As Dictionary)
Dim Ay$(): Stop ' Ay = Macro(MacroStr).Ny
If Not AyIsEmp(Ay) Then
    Dim O$: O = MacroStr
    Dim I, K$
    For Each I In Ay
        K = RmvFstLasChr(CStr(I))
        If Dic.Exists(K) Then
            O = Replace(O, I, Dic(K))
        End If
    Next
End If
FmtMacroDic = O
End Property

Property Get FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Property

Property Get FmtQQAv$(QQVbl$, Av)
If AyIsEmp(Av) Then FmtQQAv = QQVbl: Exit Property
Dim O$
    Dim I, NeedUnEsc As Boolean
    O = RplVBar(QQVbl)
    For Each I In Av
        If InStr(I, "?") > 0 Then
            NeedUnEsc = True
            I = Replace(I, "?", Chr(255))
        End If
        O = Replace(O, "?", I, Count:=1)
    Next
    If NeedUnEsc Then O = Replace(O, Chr(255), "?")
FmtQQAv = O
End Property
