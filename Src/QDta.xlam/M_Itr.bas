Attribute VB_Name = "M_Itr"
Option Explicit

Function ItrCntByBoolPrp&(A, BoolPrpNm$)
If A.Count = 0 Then Exit Function
Dim O, Cnt&
For Each O In A
    If CallByName(O, BoolPrpNm, VbGet) Then
        Cnt = Cnt + 1
    End If
Next
ItrCntByBoolPrp = Cnt
End Function

Function ItrDrs(Itr, PrpNy0) As Drs
Dim Ny$()
    Ny = DftNy(PrpNy0)
Dim Dry()
    Dim Obj
    If Itr.Count > 0 Then
        For Each Obj In Itr
            Push Dry, ObjPrpDr(Obj, Ny)
        Next
    End If
Set ItrDrs = Drs(Ny, Dry)
End Function

Function ItrItmByPrp(A, PrpNm$, PrpV)
Dim O, V
If A.Count > 0 Then
    For Each O In A
        V = CallByName(O, PrpNm, VbGet)
        If V = PrpV Then
            Asg O, ItrItmByPrp
            Exit Function
        End If
    Next
End If
End Function

Function ItrNy(A, Optional Lik$ = "*") As String()
Dim O$(), Obj, N$
If A.Count > 0 Then
    For Each Obj In A
        N = Obj.Name
        If N Like Lik Then Push O, N
    Next
End If
ItrNy = O
End Function

Private Sub ZZ_ItrDrs()
DrsDmp ItrDrs(Application.VBE.VBProjects, "Name Type")
End Sub
