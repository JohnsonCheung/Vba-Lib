Attribute VB_Name = "M_Ws"
Option Explicit
Property Get WsA1(A As Worksheet) As Range
Set WsA1 = A.Range("A1")
End Property

Property Get WsCRR(A As Worksheet, C, R1, R2) As Range
Set WsCRR = WsRCRC(A, R1, C, R2, C)
End Property

Property Get WsDftLoNm$(A As Worksheet, Optional LoNm0$)
Dim LoNm$: LoNm = DftStr(LoNm0, "Table")
Dim J%
For J = 1 To 999
    If Not WsHasLoNm(A, LoNm) Then WsDftLoNm = LoNm: Exit Property
    LoNm = NmNxtSeqNm(LoNm)
Next
Stop
End Property

Sub WsDlt(A As Workbook, WsIx)
If WbHasWs(A, WsIx) Then WbWs(A, WsIx).Delete
End Sub

Property Get WsDtaRg(A As Worksheet) As Range
Dim R, C
With WsLasCell(A)
   R = .Row
   C = .Column
End With
If R = 1 And C = 1 Then Exit Property
Set WsDtaRg = WsRCRC(A, 1, 1, R, C)
End Property

Property Get WsHasLoNm(A As Worksheet, LoNm$) As Boolean
Dim I
For Each I In A.ListObjects
    If CvLo(I).Name = LoNm Then WsHasLoNm = True: Exit Property
Next
End Property

Property Get WsLasCell(A As Worksheet) As Range
Set WsLasCell = A.Cells.SpecialCells(xlCellTypeLastCell)
End Property

Property Get WsLasCno%(A As Worksheet)
WsLasCno = WsLasCell(A).Column
End Property

Property Get WsLasRno%(A As Worksheet)
WsLasRno = WsLasCell(A).Row
End Property

Property Get WsLo(A As Worksheet, Optional LoNm$ = "Table1") As ListObject
Dim O As ListObject
Set O = A.ListObjects.Add(xlSrcRange, WsDtaRg(A), xlNo)
Dim N$: N = WsDftLoNm(A, LoNm)
If LoNm <> N Then O.Name = N
Set WsLo = O
End Property

Property Get WsNxtLoNm$(A As Worksheet, LoNm$)

End Property

Property Get WsRC(A As Worksheet, R, C) As Range
Set WsRC = A.Cells(R, C)
End Property

Property Get WsRCRC(A As Worksheet, R1, C1, R2, C2) As Range
Set WsRCRC = A.Range(WsRC(A, R1, C1), WsRC(A, R2, C2))
End Property

Property Get WsSq(A As Worksheet) As Variant()
WsSq = WsDtaRg(A).Value
End Property

Property Get WsWb(A As Worksheet) As Workbook
Set WsWb = A.Parent
End Property


Sub WsClsNoSav(A As Worksheet)
WbClsNoSav WsWb(A)
End Sub

Sub WsVis(A As Worksheet)
A.Application.Visible = True
End Sub

