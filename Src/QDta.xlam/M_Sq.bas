Attribute VB_Name = "M_Sq"
Option Explicit
Property Get SqNRow&(A)
On Error Resume Next
SqNRow = UBound(A, 1)
End Property
Property Get SqNCol&(A)
On Error Resume Next
SqNCol = UBound(A, 2)
End Property
Property Get SqTranspose(A) As Variant()
Dim NRow&, NCol&
NRow = SqNRow(A)
NCol = SqNCol(A)
Dim O(), J&, I&
ReDim O(1 - NCol, 1 To NRow)
For J = 1 To NRow
    For I = 1 To NCol
        O(I, J) = A(J, I)
    Next
Next
SqTranspose = O
End Property
Sub SqBrw(A)
DryBrw SqDry(A)
End Sub
Property Get SqRg(A, At As Range, Optional LoNm$) As Range
If Sz(A) = 0 Then Set SqRg = At.Cells(1, 1): Exit Property
Dim O As Range
Set O = RgReSz(At, A)
O.Value = A
Set SqRg = O
End Property

Property Get TitAy_Sq(TitAy$()) As Variant()
Dim UFld%: UFld = UB(TitAy)
Dim ColVBar()
    ReDim ColVBar(UFld)
    Dim J%
    For J = 0 To UFld
        ColVBar(J) = AyTrim(SplitVBar(TitAy(J)))
    Next
Dim NRow%
    Dim M%, VBar$()
    For J = 0 To UB(ColVBar)
        VBar = ColVBar(J)
        M = Sz(VBar)
        If M > NRow Then NRow = M
    Next
Dim O()
    Dim I%
    ReDim O(1 To NRow, 1 To UFld + 1)
    For J = 0 To UFld
        VBar = ColVBar(J)
        For I = 0 To UB(VBar)
            O(I + 1, J + 1) = VBar(I)
        Next
    Next

End Property
Private Sub ZZ_TitAy_Sq()
Dim A$()
Push A, "ksdf | skdfj  |skldf jf"
Push A, "skldf|sdkfl|lskdf|slkdfj"
Push A, "askdfj|sldkf"
Push A, "fskldf"
SqBrw TitAy_Sq(A)
End Sub

Property Get DrsSel(A As Drs, SelFny0, Optional AsFny0) As Drs
Dim SelFny$()
Dim AsFny$()
    SelFny = DftNy(SelFny0)
    AsFny = DftNy(AsFny0)
    If Sz(AsFny) = 0 Then
        AsFny = SelFny
    Else
        If Sz(SelFny) <> Sz(AsFny) Then Stop
    End If

Dim CIxAy%()
Dim Dry()
   Dry = A.Dry
   Dim Ay&()
   Dim U%
   Dim J%
   Stop '
'   Ay = AyIxAy(Dry(0), Fm)
   U = UB(A)
   ReDim CIxAy(U)
   For J = 0 To U
       CIxAy(J) = Ay(J)
   Next
Dim O()
    Dim R&
    For R = 0 To UBound(Dry)
        Push O, AyWhIxAy(Dry(R), CIxAy)
    Next
Set DrsSel = Drs(AsFny, O)
End Property
