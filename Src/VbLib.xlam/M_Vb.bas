Attribute VB_Name = "M_Vb"
Option Explicit
Type RRCC
    R1 As Long
    C1 As Long
    R2 As Long
    C2 As Long
End Type
Type LnoCnt
    Lno As Long
    Cnt As Long
End Type

Function DftTpLy(Tp0) As String()
Stop
'Select Case True
'Case V(Tp0).IsStr: DftTpLy = SplitVBar(Tp0)
'Case IsSy(Tp0):  DftTpLy = Tp0
'Case Else: Stop
'End Select
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

Function IsEq(Act, Exp) As Boolean
'If VarType(Act) <> VarType(Exp) Then Exit Function
'If IsPrim(Act) Then
'    If Act <> Exp Then Exit Function
'End If
'If IsArray(Act) Then
'    If Not AyIsEq(Act, Exp) Then Stop
'    Exit Function
'End If
End Function

Function CollObjAy(Coll) As Object()
Dim O() As Object
Dim V
For Each V In Coll
   Push O, V
Next
CollObjAy = O
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

Sub Never()
Const CSub$ = "Never"
Stop
'Er CSub, "Should never reach here"
End Sub

''
Sub AssChk(Chk$())
If AyIsEmp(Chk) Then Exit Sub
'AyBrw Chk
Stop
End Sub


Function Ly0Ap_Ly(ParamArray Ly0Ap()) As String()
Dim I, Av(): Av = Ly0Ap
If AyIsEmp(Av) Then Exit Function
Dim O$()
For Each I In Av
Stop
'    PushAy O, DftLy(I)
Next
Ly0Ap_Ly = O
End Function

Sub DtaEr()
MsgBox "DtaEr"
Stop
End Sub

Sub PmEr()
MsgBox "Parameter Er"
Stop
End Sub

Sub Stp()
Stop
End Sub

Function NowStr$()
NowStr = Format(Now(), "YYYY-MM-DD HH:MM:SS")
End Function

Function ObjPth(Obj, Pth$)
'Ret the Obj's Get-Property-Value using Pth, which is dot-separated-string
Dim Ny$()
    Ny = Split(Pth, ".")
Dim O
    Dim J%, U%
    Set O = Obj
    U = UB(Ny)
    For J = 0 To U - 1      ' U-1 is to skip the last Pth-Seg
        Set O = CallByName(O, Ny(J), VbGet) ' in the middle of each path-seg, they must be object, so use [Set O = ...] is OK
    Next

ObjPth = CallByName(O, Ny(U), VbGet) ' Last Prp may be non-object, so must use 'Asg'
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
If IsPrim(V) Then
   VarLy = ApSy(V)
ElseIf IsArray(V) Then
   VarLy = AySy(V)
ElseIf IsObject(V) Then
   VarLy = ApSy("*Type: " & TypeName(V))
Else
   Stop
End If
End Function

