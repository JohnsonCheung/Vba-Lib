Attribute VB_Name = "M_Fmt"
''
''Function FmtMacro$(MacroStr$, ParamArray Ap())
''Dim Av(): Av = Ap
''FmtMacro = FmtMacroAv(MacroStr, Av)
''End Function
''
''Function FmtMacroAv$(MacroStr$, Av())
''Dim Ay$(): Stop ' Ay = Macro(MacroStr).Ny
''Dim O$: O = MacroStr
''Dim J%, I
''For Each I In Ay
''    O = Replace(O, I, Av(J))
''    J = J + 1
''Next
''FmtMacroAv = O
''End Function
''
''Function FmtMacroDic$(MacroStr$, Dic As Dictionary)
''Dim Ay$(): Stop ' Ay = Macro(MacroStr).Ny
''If Not AyIsEmp(Ay) Then
''    Dim O$: O = MacroStr
''    Dim I, K$
''    For Each I In Ay
''        K = RmvFstLasChr(CStr(I))
''        If Dic.Exists(K) Then
''            O = Replace(O, I, Dic(K))
''        End If
''    Next
''End If
''FmtMacroDic = O
''End Function
''
Function FmtQQ$(QQVbl$, ParamArray Ap())
Dim Av(): Av = Ap
FmtQQ = FmtQQAv(QQVbl, Av)
End Function

Function FmtQQAv$(QQVbl$, Av)
If AyIsEmp(Av) Then FmtQQAv = QQVbl: Exit Function
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
End Function
'''
'''Function FmtQQVBar$(QQStr$, ParamArray Ap())
'''Dim Av(): Av = Ap
'''FmtQQVBar = RplVBar(FmtQQAv(QQStr, Av))
'''End Function
'''
