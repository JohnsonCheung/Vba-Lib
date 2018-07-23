Attribute VB_Name = "M_Er"
Option Explicit

Function ErMsgLines$(CSub$, MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
ErMsgLines = ErMsgLinesByAv(CSub, MacroStr, Av)
End Function

Sub Er(CSub$, MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
AyBrw ErMsgLyByAv(CSub, MacroStr, Av)
Stop
End Sub

Sub ErDta()
MsgBox "ErDta"
Stop
End Sub

Sub ErNever()
MsgBox "Should never reach here"
Stop
End Sub

Sub ErPm()
MsgBox "Parameter Er"
Stop
End Sub

Private Sub ErMsgBrw(CSub$, MacroStr$, Av())
AyBrw ErMsgLyByAv(CSub, MacroStr, Av())
End Sub

Private Function ErMsgLinesByAv$(CSub$, MacroStr$, Av())
ErMsgLinesByAv = JnCrLf(ErMsgLyByAv(CSub, MacroStr, Av))
End Function

Private Function ErMsgLyByAv(CSub$, MacroStr$, Av()) As String()
Dim O$()
   Push O, "Subr-" & CSub & " : " & RplVBar(MacroStr)
If Not AyIsEmp(Av) Then
   Dim Ny$(): Ny = MacroNy(MacroStr)
   Dim I, J%
   If Not AyIsEmp(Ny) Then
       For Each I In Ny
           Push O, Chr(9) & I
           PushAy O, AyAddPfx(VarLy(Av(J)), Chr(9) & Chr(9))
           J = J + 1
       Next
   End If
End If
ErMsgLyByAv = O
End Function
