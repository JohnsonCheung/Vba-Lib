Attribute VB_Name = "M_Vb"
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
    PushAy O, DftLy(I)
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

Property Get ObjCompoundPrp$(Obj, PrpSsl$)
Dim Ny$(): Ny = SslSy(PrpSsl)
Dim O$(), I
For Each I In Ny
    Push O, CallByName(Obj, CStr(I), VbGet)
Next
ObjCompoundPrp = Join(O, "|")
End Property

Property Get ObjPrp(Obj, PrpPth$)
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim Ny$()
    Ny = Split(PrpPth, ".")
Dim O
    Dim J%, U%
    Set O = Obj
    U = UB(Ny)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, Ny(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next

ObjPrp = CallByName(O, Ny(U), VbGet) ' Last Prp may be non-object, so must use 'Asg'
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

Property Get ToCellStr$(V, Optional ShwZer As Boolean)
'CellStr is a string can be displayed in a cell
If M_Is.IsEmp(V) Then Exit Property
If IsStr(V) Then
    ToCellStr = V
    Exit Property
End If
If IsBool(V) Then
    ToCellStr = IIf(V, "TRUE", "FALSE")
    Exit Property
End If

If IsObject(V) Then
    ToCellStr = "[" & TypeName(V) & "]"
    Exit Property
End If
If ShwZer Then
    If IsNumeric(V) Then
        If V = 0 Then ToCellStr = "0"
        Exit Property
    End If
End If
If IsArray(V) Then
    If AyIsEmp(V) Then Exit Property
    ToCellStr = "Ay" & UB(V) & ":" & V(0)
    Exit Property
End If
If InStr(V, vbCrLf) > 0 Then
    ToCellStr = Brk(V, vbCrLf).S1 & "|.."
    Exit Property
End If
ToCellStr = V
End Property

Property Get ToLy(V) As String()
If IsPrim(V) Then
   ToLy = ApSy(V)
ElseIf IsArray(V) Then
   ToLy = AySy(V)
ElseIf IsObject(V) Then
   ToLy = ApSy("*Type: " & TypeName(V))
Else
   Stop
End If
End Property

Property Get Tst() As Tst
Static Y As New Tst
Set Tst = Y
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

Sub DtaEr()
MsgBox "DtaEr"
Stop
End Sub

Sub NeverEr()
MsgBox "Should never reach here"
Stop
End Sub

Sub PmEr()
MsgBox "Parameter Er"
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

Sub Stp()
Stop
End Sub

Private Sub ZZZ_ObjCompoundPrp()
Dim Act$: Act = ObjCompoundPrp(Excel.Application.VBE.ActiveVBProject, "FileName Name")
Ass Act = "C:\Users\user\Desktop\Vba-Lib-1\QVb.xlam|QVb"
End Sub
