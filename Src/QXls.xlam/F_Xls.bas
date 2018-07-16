Attribute VB_Name = "F_Xls"
Option Explicit

Property Get NmNxtSeqNm$(A, Optional NDig% = 3)
If NDig = 0 Then Stop
Dim R$: R = Right(A, NDig + 1)
If Left(R, 1) = "_" Then
    If IsNumeric(Mid(R, 2)) Then
        Dim L$: L = Left(A, Len(A) - NDig)
        Dim Nxt%: Nxt = Val(Mid(R, 2)) + 1
        NmNxtSeqNm = L + ZerFill(Nxt, NDig)
        Exit Property
    End If
End If
NmNxtSeqNm = A & "_" & StrDup(NDig - 1, "0") & "1"
End Property

Property Get NmSeqNo%(A)
Dim B$: B = TakAftRev(A, "_")
If B = "" Then Exit Property
If Not IsNumeric(B) Then Exit Property
NmSeqNo = B
End Property

Property Get TitS1S2Ay_Sq(TitS1S2Ay() As S1S2, Fny$()) As Variant()
Dim TitColAy$():   TitColAy = A_TitColAy(TitS1S2Ay, Fny)
Dim VBarColAy():  VBarColAy = A_VBarColAy(TitColAy)
Dim NRow%:             NRow = A_TitNRow(VBarColAy)
Dim Sq(): ReDim Sq(1 To NRow, 1 To Sz(Fny))
Dim J%, C%, R%, VBar$()
For C = 0 To UB(Fny)
    VBar = VBarColAy(C)
    For R = 0 To UB(VBar)
        Sq(R + 1, C + 1) = VBar(R)
    Next
Next
TitS1S2Ay_Sq = Sq
End Property
Property Get ZerFill$(N%, NDig%)
ZerFill = Format(N, StrDup(NDig, 0))
End Property


Sub TitRg_Fmt(A As Range)
'---
    Dim C%
    For C = 1 To A.Columns.Count
        VBar_MgeBottomEmpCell RgC(A, C)
    Next
HBar_MgeSamValCell A
End Sub


Private Property Get A_TitColAy(TitS1S2Ay() As S1S2, Fny$()) As String()
Dim O$(), J%, I%, UTit%
UTit = UB(TitS1S2Ay)
For J = 0 To UB(Fny)
    For I = 0 To UTit
        If TitS1S2Ay(I).S1 = Fny(J) Then Push O, TitS1S2Ay(I).S2: GoTo Nxt
    Next
    Push O, Fny(J)
Nxt:
Next
A_TitColAy = O
End Property

Private Property Get A_TitNRow%(VBarColAy())
Dim M%, J%, S%
For J = 0 To UB(VBarColAy)
    S = Sz(VBarColAy(J))
    If S > M Then M = S
Next
A_TitNRow = M
End Property

Private Property Get A_VBarColAy(TitColAy$()) As Variant()
Dim O(), J%
For J = 0 To UB(TitColAy)
    Dim VBar$()
    VBar = AyTrim(SplitVBar(TitColAy(J)))
    Push O, VBar
Next
A_VBarColAy = O
End Property

Private Sub ZZ_TitS1S2Ay_Sq()
Dim Fny$()
    PushAy Fny, Array("X", "A", "C", "B")
Dim TitS1S2Ay() As S1S2
    PushObj TitS1S2Ay, S1S2("A", "skldf|lsjdf|sdldf")
    PushObj TitS1S2Ay, S1S2("C", "skldf|lsjdf|sdlkf |sdfsdf")
    PushObj TitS1S2Ay, S1S2("B", "skldf|ls|df|jdf|sdlkf |sdfsdf")
    PushObj TitS1S2Ay, S1S2("X", "skdf|df|lsjdf|sdlkf |sdfsdf")
'SqBrw TitS1S2Ay_Sq(TitS1S2Ay, Fny)
Stop
End Sub
