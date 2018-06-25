Attribute VB_Name = "M_Sq"
Option Explicit
Property Get SqTranspose() As Variant()
Dim NRow&, NCol&
NRow = Me.NRow
NCol = Me.NCol
Dim O(), J&, I&
ReDim O(1 - NCol, 1 To NRow)
For J = 1 To NRow
    For I = 1 To NCol
        O(I, J) = A(J, I)
    Next
Next
Transpose = Sqx(O).Sq
End Property
Sub SqBrw()
Stop '
End Sub
Property Get SqNCol%(A)
On Error Resume Next
NCol = UBound(A, 2)
End Property
Property Get SqRg(A, At As Range, Optional LoNm$) As Range
If IsEmp Then Set Rg = At.Cells(1, 1): Exit Property
Dim O As Range
Set O = RgReSz(At, Sq)
O.Value = Sq
Set Rg = O
End Property

Property Get SqNRow&(A)
On Error Resume Next
NRow = UBound(A, 1)
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
SqxByTitAy(A).Brw
End Sub

Property Get SqSel(Optional MapStr$) As Drs
Stop '
Dim Fny$(), Fm$() 'MapStr
   If MapStr = "" Then
       Fny = AySy(SqDr(Sq, 1))
       Fm = Fny
   Else
        Stop '
'       With S1S2AyStr_SyPair(MapStr)
'           Fny = .Sy1
'           Fm = .Sy2
'       End With
   End If
Dim SqCnoAy%() 'Fm,Sq
   Dim A&()
   Dim U%
   Dim J%
   A = AyIxAy(SqDr(Sq, 1), Fm)
   U = UB(A)
   ReDim SqCnoAy(U)
   For J = 0 To U
       SqCnoAy(J) = A(J) + 1
   Next
Dim Dry() 'Sq,SqIxAy
   Dim R&, Cno%, C%
   Dim UFld%
   Dim Ix%
   Dim Dr()
   UFld = UB(SqCnoAy)
   For R = 2 To UBound(Sq, 1)
       ReDim Dr(UFld)
       For C = 0 To UFld
           Cno = SqCnoAy(C)
           If Cno > 0 Then
               Dr(C) = A(R, Cno)
           End If
       Next
       Push Dry, Dr
   Next
SqSel.Dry = Dry
SqSel.Fny = Fny
End Property


