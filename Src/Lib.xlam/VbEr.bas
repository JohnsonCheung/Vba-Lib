Attribute VB_Name = "VbEr"
Option Explicit
Public Fso As New FileSystemObject

Sub Er(CSub$, MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
AyBrw ErMsgLyByAv(CSub, MacroStr, Av)
Stop
End Sub

Sub ErMsgBrw(CSub$, MacroStr$, Av())
AyBrw ErMsgLyByAv(CSub, MacroStr, Av())
End Sub
Function ErMsgLines$(CSub$, MacroStr$, ParamArray Ap())
Dim Av(): Av = Ap
ErMsgLines = ErMsgLinesByAv(CSub, MacroStr, Av)
End Function

Function ErMsgLinesByAv$(CSub$, MacroStr$, Av())
ErMsgLinesByAv = JnCrLf(ErMsgLyByAv(CSub, MacroStr, Av))
End Function

Function ErMsgLyByAv(CSub$, MacroStr$, Av()) As String()
Dim O$()
   Push O, "Subr-" & CSub & " : " & RplVBar(MacroStr)
If Not AyIsEmp(Av) Then
   Dim Ny$(): Ny = MacroStr_Ny(MacroStr)
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

Function NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Function

Function ObjPrp(Obj, PrpNm$)
ObjPrp = CallByName(Obj, PrpNm, VbGet)
End Function

Function PipeAy(Prm, MthNy$())
Dim O: Asg Prm, O
Dim I
For Each I In MthNy
   Asg Run(I, O), O
Next
Asg O, PipeAy
End Function

Function RunAv(MthNm$, Av)
Dim O
Select Case Sz(Av)
Case 0: O = Run(MthNm)
Case 1: O = Run(MthNm, Av(0))
Case 2: O = Run(MthNm, Av(0), Av(1))
Case 3: O = Run(MthNm, Av(0), Av(1), Av(2))
Case 4: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3))
Case 5: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4))
Case 6: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5))
Case 7: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6))
Case 8: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7))
Case 9: O = Run(MthNm, Av(0), Av(1), Av(2), Av(3), Av(4), Av(5), Av(6), Av(7), Av(8))
Case Else: Stop
End Select
RunAv = O
End Function

Function VarLy(V) As String()
If ValIsPrim(V) Then
   VarLy = ApSy(V)
ElseIf IsArray(V) Then
   VarLy = AySy(V)
ElseIf IsObject(V) Then
   VarLy = ApSy("*Type: " & TypeName(V))
Else
   Stop
End If
End Function

Sub Tst()
End Sub
