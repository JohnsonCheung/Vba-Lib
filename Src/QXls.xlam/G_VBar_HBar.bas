Attribute VB_Name = "G_VBar_HBar"
Option Explicit

Property Get HBar_SamValRg(A As Range) As Range
Dim NCol%: NCol = RgNCol(A)
Dim C1%, V, C2%, Fnd As Boolean
For C1 = 1 To NCol - 1
    V = RgRC(A, 1, C1).Value
    For C2 = C1 + 1 To NCol
        If RgRC(A, 1, C2).Value = V Then
            Fnd = True
        Else
            If Fnd Then
                C2 = C2 - 1
                GoTo Fnd
            End If
            GoTo Nxt
        End If
    Next
Nxt:
Next
Fnd:
If Fnd Then Set HBar_SamValRg = RgRCC(A, 1, C1, C2)
End Property

Property Get VBarAy(A As Range) As Variant()
Ass RgIsVBar(A)
VBarAy = SqCol(RgSq(A), 1)
End Property

Property Get VBarIntAy(A As Range) As Integer()
VBarIntAy = AyIntAy(VBarAy(A))
End Property

Sub HBar_MgeSamValCell(A As Range)
Ass RgIsHBar(A)
Dim R As Range
Set R = HBar_SamValRg(A)
Dim Sav As Boolean
    Sav = A.Application.DisplayAlerts
    A.Application.DisplayAlerts = False
While Not IsNothing(R)
    R.Merge '<===================================
    Set R = HBar_SamValRg(R)
Wend
A.Application.DisplayAlerts = Sav
End Sub


Property Get VBarSy(A As Range) As String()
VBarSy = AySy(VBarAy(A))
End Property


Sub VBar_MgeBottomEmpCell(A As Range)
Ass RgIsVBar(A)
Dim R2: R2 = A.Rows.Count
Dim R1
    Dim Fnd As Boolean
    For R1 = R2 To 1 Step -1
        If Not IsEmpty(RgRC(A, R1, 1)) Then Fnd = True: GoTo Nxt
    Next
Nxt:
    If Not Fnd Then Stop
If R2 = R1 Then Exit Sub
Dim R As Range: Set R = RgCRR(A, 1, R1, R2)
R.Merge
R.VerticalAlignment = XlVAlign.xlVAlignTop
End Sub
