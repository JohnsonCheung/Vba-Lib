Attribute VB_Name = "DtaSq"
Option Explicit
Function SqRg(A, At As Range) As Range
If Sz(A) = 0 Then Exit Function
Dim O As Range: Set O = CellReSz(At, A)
O.Value = A
Set SqRg = O
End Function
Function SqLo(A, At As Range, Optional LoNm$) As ListObject
Set SqLo = RgLo(SqRg(A, At), LoNm)
End Function
Function SqWs(A, Optional WsNm$ = "Sheet1") As Worksheet
Dim A1 As Range: Set A1 = NewA1
SqRg A, A1
Set SqWs = RgWs(A1)
End Function
Function SqRow(A, R%) As String()
Dim J%
For J = 1 To UBound(A, 2)
    Push SqRow, A(R, J)
Next
End Function
Function SqLy(A) As String()
Dim R%
For R = 1 To UBound(A, 1)
    Push SqLy, JnSpc(SqRow(A, R))
Next
End Function
Function SqAlign(Sq(), W%()) As Variant()
If UBound(Sq, 2) <> Sz(W) Then Stop
Dim C%, R%, Wdt%, O
O = Sq
For C = 1 To UBound(Sq, 2) - 1 ' The last column no need to align
    Wdt = W(C - 1)
    For R = 1 To UBound(Sq, 1)
        O(R, C) = AlignL(Sq(R, C), Wdt)
    Next
Next
SqAlign = O
End Function
Sub SqSetRow(OSq, R&, Dr)
Dim J%
For J = 0 To UB(Dr)
    OSq(R, J + 1) = Dr(J)
Next
End Sub
Function SqBktEndPos%(A, Optional FmPos% = 1)
SqBktEndPos = BktXEndPos(A, "[", "]", FmPos)
End Function
