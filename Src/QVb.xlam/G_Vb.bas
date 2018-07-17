Attribute VB_Name = "G_Vb"
Option Explicit

Property Get DftTpLy(Tp0) As String()
Stop
'Select Case True
'Case V(Tp0).IsStr: DftTpLy = SplitVBar(Tp0)
'Case IsSy(Tp0):  DftTpLy = Tp0
'Case Else: Stop
'End Select
End Property

Property Get Fso() As Scripting.FileSystemObject
Static Y As New Scripting.FileSystemObject
Set Fso = Y
End Property

Property Get IsEq(Act, Exp) As Boolean
'If VarType(Act) <> VarType(Exp) Then Exit Function
'If IsPrim(Act) Then
'    If Act <> Exp Then Exit Function
'End If
'If IsArray(Act) Then
'    If Not AyIsEq(Act, Exp) Then Stop
'    Exit Function
'End If
End Property

Property Get Ly0Ap_Ly(ParamArray Ly0Ap()) As String()
Dim I, Av(): Av = Ly0Ap
If AyIsEmp(Av) Then Exit Property
Dim O$()
For Each I In Av
    PushAy O, DftNy(I)
Next
Ly0Ap_Ly = O
End Property

Property Get Max(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) > O Then O = Av(J)
Next
Max = O
End Property

Property Get Min(ParamArray Ap())
Dim Av(), O
Av = Ap
O = Av(0)
Dim J%
For J = 1 To UB(Av)
   If Av(J) < O Then O = Av(J)
Next
Min = O
End Property

Property Get NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Property


Property Get PgmPth$()
PgmPth = FfnPth(Excel.Application.VBE.ActiveVBProject.Filename)
End Property

Property Get Pipe(Pm, MthNy0)
Dim O: Asg Pm, O
Dim I
For Each I In DftNy(MthNy0)
   Asg Run(I, O), O
Next
Asg O, Pipe
End Property

Property Get PipeAy(Prm, MthNy$())
Dim O: Asg Prm, O
Dim I
For Each I In MthNy
   Asg Run(I, O), O
Next
Asg O, PipeAy
End Property

Property Get Tst() As Tst
Static Y As New Tst
Set Tst = Y
End Property

Property Get ZerFill$(N%, NDig%)
ZerFill = Format(N, StrDup(NDig, 0))
End Property

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
