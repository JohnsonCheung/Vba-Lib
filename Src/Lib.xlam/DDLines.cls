VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DDLines"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private A As Dictionary

Sub Brw()
AyBrw FmtLy
End Sub
Friend Sub Tst()

End Sub
Function Dic2(K1) As Dictionary
Dim mK1$: mK1 = K1
If A.Exists(mK1) Then
    Set Dic2 = A(mK1)
End If
End Function

Function FmtLy() As String()
Dim O$(), K1, D2 As Dictionary
For Each K1 In A
    Set D2 = Dic2(K1)
    PushAy O, AyAddPfx(DicLy(D2, Opt:=e_KeyLinesFmt), K1 & " ")
Next
FmtLy = O
End Function

Property Get Init(Ly$()) As DDLines
Dim J&
For J = 0 To UB(Ly)
    AddLin Ly(J)
Next
Set Init = Me
End Property

Function K1_S1S2Ay(K1) As S1S2()
Dim Ky2$(): Ky2 = Me.Ky2(K1)
Dim O() As S1S2, J%
For J = 0 To UB(Ky2)
    S1S2_Push O, NewS1S2(Ky2(J), KKLines(K1, Ky2(J)))
Next
K1_S1S2Ay = O
End Function

Function KKLines$(K1, K2)
Dim D2 As Dictionary: Set D2 = Dic2(K1): If IsNothing(D2) Then Exit Function
KKLines = DicVal(D2, CStr(K2))
End Function

Function Ky2(K1) As String()
Dim D2 As Dictionary: Set D2 = Dic2(K1)
If IsNothing(D2) Then Exit Function
Ky2 = DicKeySy(D2)
End Function

Private Sub AddKKLin(K1$, K2$, Lin$)
If FstChr(Lin) = "~" Then Lin = RplFstChr(Lin, " ")
Dim D2 As Dictionary, L$
If A.Exists(K1) Then
    Set D2 = A(K1)
    If D2.Exists(K2) Then
        L = D2(K2)
        If L = "" Then
            L = Lin
        Else
            L = L & vbCrLf & Lin
        End If
        D2(K2) = L
    Else
        D2.Add K2, Lin
    End If
Else
    Set D2 = NewDic
    D2.Add K2, Lin
    A.Add K1, D2
End If
End Sub

Private Sub AddLin(L)
Dim K1$, K2$, Ln$: Lin(L).AsgTTRst K1, K2, Ln
AddKKLin K1, K2, Ln
End Sub

Private Sub Class_Initialize()
Set A = New Dictionary
End Sub

Private Sub Class_Terminate()
On Error Resume Next
A.RemoveAll
Set A = Nothing
Debug.Print "DDLines: Terminate ..."
End Sub

Private Sub ZZResSampleDDLinesLy()
'A B lskdf
'A B ~ lsdkf
'A C sdkfjdf
'A C ~ sdfsdf
'B X sldkf sdf
End Sub

Private Function ZZSampleDDLines() As DDLines
Set ZZSampleDDLines = NewDDLines(ZZSampleDDLinesLy)
End Function

Private Function ZZSampleDDLinesLy() As String()
'ZZSampleDDLinesLy = MdResLy(Md("DDLines"), "SampleDDLinesLy")
End Function

Private Sub FmtLy__Tst()
Dim D As DDLines: Set D = ZZSampleDDLines
AyBrw D.FmtLy
End Sub

Sub KKLines__Tst()
With ZZSampleDDLines
    Ass .KKLines("B", "X") = "sldkf sdf"
    Ass .KKLines("A", "B") = RplVBar("lskdf|  lsdkf")
    Ass .KKLines("A", "C") = RplVBar("sdkfjdf|  sdfsdf")
End With
End Sub

Private Sub Ky2__Tst()
Dim A$()
With ZZSampleDDLines
    A = .Ky2("A")
    Ass Sz(A) = 2
    Ass A(0) = "B"
    Ass A(1) = "C"
    
    A = .Ky2("B")
    Ass Sz(A) = 1
    Ass A(0) = "X"
    
    A = .Ky2("C")
    Ass Sz(A) = 0
End With
End Sub

Private Sub ZZSamplesDDLinesLy__Tst()
AyBrw ZZSampleDDLinesLy
End Sub

