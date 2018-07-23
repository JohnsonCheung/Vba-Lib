Attribute VB_Name = "G_Vb"
Option Explicit

Function Fso() As Scripting.FileSystemObject
Static Y As New Scripting.FileSystemObject
Set Fso = Y
End Function

Function Ly0Ap_Ly(ParamArray Ly0Ap()) As String()
Dim I, Av(): Av = Ly0Ap
If AyIsEmp(Av) Then Exit Function
Dim O$()
For Each I In Av
    PushAy O, DftNy(I)
Next
Ly0Ap_Ly = O
End Function

Function Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) > O Then O = Av(J)
Next
Max = O
End Function

Function Min(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) < O Then O = Av(J)
Next
Min = O
End Function

Function NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Function

Function PgmPth$()
PgmPth = FfnPth(Excel.Application.VBE.ActiveVBProject.Filename)
End Function

Function Pipe(Pm, MthNy0)
Dim O: Asg Pm, O
Dim I
For Each I In DftNy(MthNy0)
   Asg Run(I, O), O
Next
Asg O, Pipe
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

Function Tst() As Tst
Static Y As New Tst
Set Tst = Y
End Function

Function ZerFill$(N%, NDig%)
ZerFill = Format(N, StrDup(NDig, 0))
End Function

Sub Asg(V, OV)
If IsObject(V) Then
   Set OV = V
Else
   OV = V
End If
End Sub

Sub Ass(A As Boolean)
Debug.Assert A
End Sub

''
Sub Chk(Check$())
If AyIsEmp(Check) Then Exit Sub
AyBrw Check
Stop
End Sub
